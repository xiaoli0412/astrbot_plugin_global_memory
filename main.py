"""全局记忆桥 — Global Memory Bridge for AstrBot.

补齐 LivingMemory 覆盖不到的两类信息：
  1. 文件：收到 docx/xlsx/pptx/pdf/txt/md/csv 等时，抽取正文入库，可跨会话检索
  2. 事件：工具调用（谁在什么时候用什么工具做了什么）入库

不接管会话上下文、不改 system_prompt，只以临时内容注入相关片段。
"""

import hashlib
import json
import os
import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path

from astrbot.api import AstrBotConfig, logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.message_components import File, Image, Plain, Record, Reply, Video
from astrbot.api.provider import ProviderRequest
from astrbot.api.star import Context, Star, register
from astrbot.core.agent.message import TextPart

PLUGIN_NAME = "astrbot_plugin_global_memory"
PLUGIN_VERSION = "0.1.0"
DATA_DIR = Path("/AstrBot/data/plugin_data") / PLUGIN_NAME
DB_PATH = DATA_DIR / "global_memory.db"
INJECT_LOG = DATA_DIR / "inject.log"
WORKSPACES = Path("/AstrBot/data/workspaces")
TEMP_DIR = Path("/AstrBot/data/temp")

DOC_EXTS = {
    ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".pdf",
    ".txt", ".md", ".csv", ".json", ".log", ".yaml", ".yml", ".html", ".htm",
}
MAX_EXTRACT_CHARS = 6000


# ────────────────────────── 文档正文抽取 ──────────────────────────

def _clip(text: str, limit: int) -> str:
    text = re.sub(r"[ \t]+", " ", text or "")
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if len(text) > limit:
        return text[:limit] + f"\n…（已截断，原文约 {len(text)} 字）"
    return text


def _read_docx(path: Path) -> str:
    from docx import Document

    doc = Document(str(path))
    parts = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text and c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _read_xlsx(path: Path) -> str:
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"[工作表] {ws.title}")
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 300:
                out.append("…（行数过多已截断）")
                break
            cells = [str(c) for c in row if c is not None and str(c).strip()]
            if cells:
                out.append(" | ".join(cells))
    wb.close()
    return "\n".join(out)


def _read_pptx(path: Path) -> str:
    from pptx import Presentation

    prs = Presentation(str(path))
    out = []
    for idx, slide in enumerate(prs.slides, 1):
        out.append(f"[第 {idx} 页]")
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text and shape.text.strip():
                out.append(shape.text.strip())
    return "\n".join(out)


def _read_pdf(path: Path) -> str:
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(path))
        out = []
        for i, page in enumerate(reader.pages):
            if i >= 60:
                out.append("…（页数过多已截断）")
                break
            txt = page.extract_text() or ""
            if txt.strip():
                out.append(txt)
        text = "\n".join(out)
        if text.strip():
            return text
    except Exception:
        pass
    try:
        import pdfplumber

        with pdfplumber.open(str(path)) as pdf:
            return "\n".join(
                (p.extract_text() or "") for p in pdf.pages[:60]
            )
    except Exception:
        return ""


def _read_plain(path: Path) -> str:
    raw = path.read_bytes()[: 400_000]
    try:
        import chardet

        enc = (chardet.detect(raw).get("encoding") or "utf-8")
    except Exception:
        enc = "utf-8"
    for e in (enc, "utf-8", "gb18030", "latin-1"):
        try:
            return raw.decode(e)
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace")


def read_document(path: Path) -> str:
    ext = path.suffix.lower()
    try:
        if ext == ".docx":
            return _read_docx(path)
        if ext in (".xlsx", ".xls"):
            return _read_xlsx(path)
        if ext == ".pptx":
            return _read_pptx(path)
        if ext == ".pdf":
            return _read_pdf(path)
        if ext in (".txt", ".md", ".csv", ".json", ".log", ".yaml", ".yml", ".html", ".htm"):
            return _read_plain(path)
    except Exception as exc:  # noqa: BLE001
        logger.debug(f"[{PLUGIN_NAME}] 解析 {path.name} 失败: {exc}")
    return ""


def locate_file(name: str, raw: str = "") -> Path | None:
    """在容器内定位收到的文件实体。"""
    candidates: list[Path] = []
    if raw:
        if raw.startswith("file://"):
            candidates.append(Path(raw[7:]))
        elif os.path.isabs(raw):
            candidates.append(Path(raw))
    if name:
        safe = os.path.basename(name)
        candidates.append(TEMP_DIR / safe)
        try:
            for p in TEMP_DIR.glob(f"fileseg_*{safe}*"):
                candidates.append(p)
        except Exception:
            pass
        try:
            for p in WORKSPACES.rglob(safe):
                candidates.append(p)
        except Exception:
            pass
    for c in candidates:
        try:
            if c.is_file() and c.stat().st_size > 0:
                return c
        except Exception:
            continue
    return None


# ────────────────────────── 存储 ──────────────────────────

def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class GlobalMemoryStore:
    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.path), timeout=10)
        conn.execute("PRAGMA journal_mode=WAL")
        return conn

    def _init_db(self) -> None:
        with self._conn() as c:
            c.executescript(
                """
                CREATE TABLE IF NOT EXISTS gm_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL,
                    kind TEXT NOT NULL,
                    session_id TEXT DEFAULT '',
                    platform TEXT DEFAULT '',
                    sender_id TEXT DEFAULT '',
                    sender_name TEXT DEFAULT '',
                    title TEXT DEFAULT '',
                    content TEXT DEFAULT '',
                    meta TEXT DEFAULT '{}',
                    fingerprint TEXT UNIQUE
                );
                CREATE INDEX IF NOT EXISTS idx_gm_kind ON gm_items(kind);
                CREATE INDEX IF NOT EXISTS idx_gm_time ON gm_items(created_at);
                CREATE VIRTUAL TABLE IF NOT EXISTS gm_fts USING fts5(
                    title, content, content='gm_items', content_rowid='id',
                    tokenize='trigram'
                );
                CREATE TRIGGER IF NOT EXISTS gm_ai AFTER INSERT ON gm_items BEGIN
                    INSERT INTO gm_fts(rowid, title, content)
                    VALUES (new.id, new.title, new.content);
                END;
                CREATE TRIGGER IF NOT EXISTS gm_ad AFTER DELETE ON gm_items BEGIN
                    INSERT INTO gm_fts(gm_fts, rowid, title, content)
                    VALUES ('delete', old.id, old.title, old.content);
                END;
                """
            )

    def add(self, *, kind: str, title: str, content: str, session_id: str = "",
            platform: str = "", sender_id: str = "", sender_name: str = "",
            meta: dict | None = None) -> bool:
        content = content or ""
        fp_src = f"{kind}|{session_id}|{title}|{content[:300]}|{len(content)}"
        fp = hashlib.sha256(fp_src.encode("utf-8", "ignore")).hexdigest()
        try:
            with self._conn() as c:
                c.execute(
                    """INSERT INTO gm_items
                       (created_at, kind, session_id, platform, sender_id, sender_name,
                        title, content, meta, fingerprint)
                       VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (_now(), kind, session_id, platform, sender_id, sender_name,
                     title, content, json.dumps(meta or {}, ensure_ascii=False), fp),
                )
            return True
        except sqlite3.IntegrityError:
            return False
        except Exception as exc:  # noqa: BLE001
            logger.warning(f"[{PLUGIN_NAME}] 写入失败: {exc}")
            return False

    @staticmethod
    def _ngrams(text: str) -> set[str]:
        """把中文句子拆成 2~4 字片段，英文/数字按词切，用于模糊相关性打分。"""
        out: set[str] = set()
        for run in re.findall(r"[\u4e00-\u9fff]+", text or ""):
            for n in (2, 3, 4):
                for i in range(len(run) - n + 1):
                    out.add(run[i:i + n])
        for w in re.findall(r"[A-Za-z0-9_]{2,}", text or ""):
            out.add(w.lower())
        return out

    def search(self, query: str, limit: int = 5) -> list[dict]:
        q = (query or "").strip()
        if not q:
            return []
        rows: list[tuple] = []
        try:
            with self._conn() as c:
                # 1) 短查询：先用 FTS5 trigram 精确匹配
                if 3 <= len(q) <= 12:
                    try:
                        cur = c.execute(
                            """SELECT i.id, i.created_at, i.kind, i.title, i.content,
                                      i.sender_name, i.session_id
                               FROM gm_fts f JOIN gm_items i ON i.id = f.rowid
                               WHERE gm_fts MATCH ? ORDER BY rank LIMIT ?""",
                            ("\"" + q.replace('"', " ") + "\"", limit),
                        )
                        rows = cur.fetchall()
                    except Exception:
                        rows = []
                # 2) 短语 LIKE 兜底
                if not rows:
                    like = f"%{q}%"
                    cur = c.execute(
                        """SELECT id, created_at, kind, title, content, sender_name, session_id
                           FROM gm_items
                           WHERE title LIKE ? OR content LIKE ?
                           ORDER BY id DESC LIMIT ?""",
                        (like, like, limit),
                    )
                    rows = cur.fetchall()
                # 3) 长句/自然语言：关键词片段打分
                if not rows:
                    terms = self._ngrams(q)
                    if terms:
                        cur = c.execute(
                            """SELECT id, created_at, kind, title, content, sender_name, session_id
                               FROM gm_items ORDER BY id DESC LIMIT 2000"""
                        )
                        scored = []
                        for r in cur.fetchall():
                            hay = (r[3] or "") + "\n" + (r[4] or "")
                            hay_l = hay.lower()
                            score = sum(len(t) for t in terms if t in hay_l)
                            if score >= 4:
                                scored.append((score, r))
                        scored.sort(key=lambda x: (-x[0], -x[1][0]))
                        rows = [r for _, r in scored[:limit]]
        except Exception as exc:  # noqa: BLE001
            logger.warning(f"[{PLUGIN_NAME}] 检索失败: {exc}")
            return []
        return [
            {
                "id": r[0], "created_at": r[1], "kind": r[2], "title": r[3],
                "content": r[4], "sender_name": r[5], "session_id": r[6],
            }
            for r in rows
        ]

    def recent(self, limit: int = 10) -> list[dict]:
        with self._conn() as c:
            cur = c.execute(
                """SELECT id, created_at, kind, title, substr(content,1,120), sender_name
                   FROM gm_items ORDER BY id DESC LIMIT ?""",
                (limit,),
            )
            return [
                {"id": r[0], "created_at": r[1], "kind": r[2], "title": r[3],
                 "preview": r[4], "sender_name": r[5]}
                for r in cur.fetchall()
            ]

    def stats(self) -> dict:
        with self._conn() as c:
            total = c.execute("SELECT COUNT(*) FROM gm_items").fetchone()[0]
            by_kind = dict(c.execute(
                "SELECT kind, COUNT(*) FROM gm_items GROUP BY kind"
            ).fetchall())
            last = c.execute("SELECT MAX(created_at) FROM gm_items").fetchone()[0]
        return {"total": total, "by_kind": by_kind, "last": last}


# ────────────────────────── 插件 ──────────────────────────

@register(PLUGIN_NAME, "xiaoli0412", "全局记忆桥：文件正文与工具事件入库并跨会话注入", PLUGIN_VERSION)
class GlobalMemory(Star):
    def __init__(self, context: Context, config: AstrBotConfig | None = None):
        super().__init__(context)
        self.config = config if config is not None else {}
        self.store = GlobalMemoryStore(DB_PATH)

    # ---------- config ----------
    def _cfg(self, key: str, default):
        v = self.config.get(key, default)
        return default if v is None else v

    async def initialize(self):
        s = self.store.stats()
        logger.info(
            "[全局记忆] 已就绪 | 条目=%s | 分类=%s | 注入=%s",
            s["total"], s["by_kind"] or "空",
            "开启" if self._cfg("inject_enabled", True) else "关闭",
        )
        # PATCHED(local): 注册 WebUI 记忆浏览页所需的原生 API。
        # 路由必须带插件名前缀，AstrBot 按 "<metadata.name>/<endpoint>" 匹配。
        try:
            reg = getattr(self.context, "register_web_api", None)
            if reg:
                reg(
                    "astrbot_plugin_global_memory/memories",
                    self.api_memories,
                    ["GET"],
                    "全局记忆：列出 / 搜索",
                )
                reg(
                    "astrbot_plugin_global_memory/stats",
                    self.api_stats,
                    ["GET"],
                    "全局记忆：统计",
                )
                reg(
                    "astrbot_plugin_global_memory/delete",
                    self.api_delete,
                    ["POST"],
                    "全局记忆：删除一条",
                )
                logger.info("[全局记忆] WebUI API 已注册 3 条")
        except Exception as exc:  # noqa: BLE001
            logger.warning("[全局记忆] WebUI API 注册失败：%s", exc)

    # ---------- WebUI API（PATCHED local） ----------
    async def api_memories(self):
        """列出或搜索记忆。GET ?q=&limit="""
        from astrbot.api.web import request  # type: ignore
        q = (request.query.get("q", "") or "").strip()
        try:
            limit = int(request.query.get("limit", 50))
        except Exception:
            limit = 50
        limit = max(1, min(limit, 200))
        rows = self.store.search(q, limit) if q else self.store.recent(limit)
        out = []
        for r in rows:
            out.append({
                "id": r.get("id"),
                "created_at": r.get("created_at"),
                "kind": r.get("kind"),
                "title": r.get("title"),
                "content": (r.get("content") or "")[:4000],
                "sender_name": r.get("sender_name"),
                "session_id": r.get("session_id"),
            })
        return {"ok": True, "data": {"total": len(out), "items": out}}

    async def api_stats(self):
        return {"ok": True, "data": self.store.stats()}

    async def api_delete(self):
        from astrbot.api.web import request  # type: ignore
        try:
            mid = int(request.json.get("id"))
        except Exception:
            return {"ok": False, "error": "缺少 id"}
        try:
            with self.store._conn() as c:  # noqa: SLF001
                c.execute("DELETE FROM gm_items WHERE id=?", (mid,))
            return {"ok": True, "data": {"deleted": mid}}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}

    # ---------- 捕获：文件 ----------
    async def _capture_files(self, event: AstrMessageEvent) -> None:
        if not self._cfg("capture_files", True):
            return
        try:
            chain = event.get_messages() or []
        except Exception:
            return
        seen_names = set()
        for comp in chain:
            if not isinstance(comp, File):
                continue
            seen_names.add(comp.name or comp.file or "")
            name = comp.name or os.path.basename(comp.file or "") or "未命名文件"
            path = locate_file(name, comp.file or "")
            size = path.stat().st_size if path else 0
            ext = Path(name).suffix.lower()
            text = ""
            if path and ext in DOC_EXTS and size <= int(self._cfg("max_file_mb", 30)) * 1024 * 1024:
                text = _clip(read_document(path), int(self._cfg("file_text_max_chars", MAX_EXTRACT_CHARS)))
            meta = {"file": name, "size": size, "ext": ext, "path": str(path) if path else ""}
            title = f"文件《{name}》"
            content = text or f"（仅记录文件元信息，未能提取正文；类型 {ext or '未知'}，大小 {size} 字节）"
            ok = self.store.add(
                kind="file", title=title, content=content,
                session_id=event.unified_msg_origin,
                platform=event.get_platform_name() if hasattr(event, "get_platform_name") else "",
                sender_id=str(event.get_sender_id() or ""),
                sender_name=event.get_sender_name() or "",
                meta=meta,
            )
            if ok:
                logger.info(f"[全局记忆] 已归档文件《{name}》，正文 {len(text)} 字")

    # ---------- 捕获：工具事件 ----------
    @filter.on_llm_tool_respond()
    async def capture_tool(self, event: AstrMessageEvent, tool, tool_args, tool_result):
        if not self._cfg("capture_tools", True):
            return
        try:
            name = getattr(tool, "name", None) or getattr(tool, "func_name", None) or str(tool)
            args = tool_args if isinstance(tool_args, str) else json.dumps(tool_args, ensure_ascii=False, default=str)
            result = tool_result
            if not isinstance(result, str):
                result = json.dumps(result, ensure_ascii=False, default=str)
            self.store.add(
                kind="tool",
                title=f"工具 {name}",
                content=_clip(f"参数：{args}\n结果：{result}", 1200),
                session_id=event.unified_msg_origin,
                platform=event.get_platform_name() if hasattr(event, "get_platform_name") else "",
                sender_id=str(event.get_sender_id() or ""),
                sender_name=event.get_sender_name() or "",
                meta={"tool": name},
            )
        except Exception as exc:  # noqa: BLE001
            logger.debug(f"[{PLUGIN_NAME}] 工具捕获失败: {exc}")

    # ---------- 捕获：消息入口（早于 office_assistant 改写） ----------
    @filter.event_message_type(filter.EventMessageType.ALL, priority=150)
    async def capture_message(self, event: AstrMessageEvent):
        """在消息进入管线时立即捕获文件组件，避免被下游插件改写后丢失。"""
        if not self._cfg("enabled", True):
            return
        try:
            await self._capture_files(event)
        except Exception as exc:  # noqa: BLE001
            logger.debug(f"[{PLUGIN_NAME}] 消息级捕获失败: {exc}")

    # ---------- 捕获 + 注入 ----------
    @filter.on_llm_request(priority=20)
    async def on_request(self, event: AstrMessageEvent, req: ProviderRequest):
        await self._capture_files(event)

        if not self._cfg("inject_enabled", True):
            return
        query = ""
        try:
            query = event.message_str or ""
        except Exception:
            query = ""
        if len(query.strip()) < 2:
            return

        hits = self.store.search(query, int(self._cfg("inject_top_k", 3)))
        if not hits:
            return

        max_chars = int(self._cfg("inject_max_chars", 1200))
        lines = ["<global_memory source=\"全局记忆·文件与事件档案\">"]
        lines.append("以下是与当前话题相关的历史文件/事件记录（来自其它会话或更早时间，仅供参考，不要假装是刚发生的）：")
        used = 0
        for h in hits:
            snippet = h["content"][: max(200, max_chars // max(1, len(hits)))]
            block = f"[{h['created_at']}] {h['title']}\n{snippet}"
            if used + len(block) > max_chars:
                break
            lines.append(block)
            used += len(block)
        lines.append("</global_memory>")

        text_block = "\n\n".join(lines)
        part = TextPart(text=text_block).mark_as_temp()
        try:
            with open(INJECT_LOG, "a", encoding="utf-8") as f:
                f.write(f"\n----- {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} -----\n{text_block}\n")
        except Exception:
            pass
        parts = getattr(req, "extra_user_content_parts", None)
        if parts is None:
            parts = []
            req.extra_user_content_parts = parts
        parts.append(part)

    # ---------- 工具 ----------
    @filter.llm_tool(name="global_memory_search")
    async def tool_search(self, event: AstrMessageEvent, query: str, limit: int = 5) -> str:
        """在全局记忆中检索历史文件正文与工具事件。当用户提到"之前那个文件/上次那个文档/你之前处理过"时调用。

        Args:
            query (string): 检索关键词，尽量用名词短语
            limit (number): 返回条数，默认 5
        """
        hits = self.store.search(query, max(1, min(int(limit or 5), 10)))
        if not hits:
            return "全局记忆中没有找到相关记录。"
        out = []
        for h in hits:
            out.append(f"[{h['created_at']}] {h['title']}\n{h['content'][:600]}")
        return "\n\n".join(out)

    @filter.llm_tool(name="global_memory_note")
    async def tool_note(self, event: AstrMessageEvent, title: str, content: str) -> str:
        """把一条值得长期记住的事实/事件写入全局记忆（跨会话可检索）。

        Args:
            title (string): 简短标题
            content (string): 要记住的内容
        """
        ok = self.store.add(
            kind="note", title=title or "随手记", content=content or "",
            session_id=event.unified_msg_origin,
            platform=event.get_platform_name() if hasattr(event, "get_platform_name") else "",
            sender_id=str(event.get_sender_id() or ""),
            sender_name=event.get_sender_name() or "",
        )
        return "已写入全局记忆。" if ok else "内容重复，已跳过。"

    # ---------- 指令 ----------
    @filter.command("全局记忆")
    async def cmd_memory(self, event: AstrMessageEvent):
        """查看全局记忆状态；/全局记忆 搜索 关键词；/全局记忆 最近"""
        event.stop_event()
        args = (event.message_str or "").strip().split(maxsplit=2)
        sub = args[1] if len(args) > 1 else ""
        if sub == "搜索" and len(args) > 2:
            hits = self.store.search(args[2], 5)
            if not hits:
                yield event.plain_result(f"没有找到「{args[2]}」相关记录。")
                return
            lines = [f"🔍 「{args[2]}」命中 {len(hits)} 条："]
            for h in hits:
                lines.append(f"\n【{h['created_at']}】{h['title']}\n{h['content'][:200]}")
            yield event.plain_result("\n".join(lines))
            return
        if sub == "最近":
            rows = self.store.recent(10)
            if not rows:
                yield event.plain_result("全局记忆还是空的。")
                return
            lines = ["📚 最近 10 条全局记忆："]
            for r in rows:
                lines.append(f"\n[{r['created_at']}] ({r['kind']}) {r['title']}\n{(r['preview'] or '')[:80]}")
            yield event.plain_result("\n".join(lines))
            return
        s = self.store.stats()
        yield event.plain_result(
            f"🧠 全局记忆桥 v{PLUGIN_VERSION}\n"
            f"总条目：{s['total']}\n"
            f"分类：{json.dumps(s['by_kind'], ensure_ascii=False) if s['by_kind'] else '（空）'}\n"
            f"最近写入：{s['last'] or '—'}\n"
            f"文件捕获：{'开' if self._cfg('capture_files', True) else '关'} | "
            f"工具捕获：{'开' if self._cfg('capture_tools', True) else '关'} | "
            f"自动注入：{'开' if self._cfg('inject_enabled', True) else '关'}\n"
            f"命令：/全局记忆 搜索 <关键词>、/全局记忆 最近"
        )