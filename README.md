# 全局记忆桥

> 补齐记忆插件对文件正文与工具事件的盲区，跨会话检索

## 这是什么

补齐 LivingMemory 的文件与事件盲区：文件正文抽取入库、工具调用记录、跨会话检索注入。

## 安装

```bash
cd AstrBot/data/plugins
git clone https://github.com/xiaoli0412/astrbot_plugin_global_memory.git
```

然后在 WebUI 插件管理里启用，或重启 AstrBot。

## 配置

全部配置都能在 **WebUI → 插件 → 全局记忆桥 → 配置** 里改，不用手改文件。

| 键 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `enabled` | bool | `True` | 关闭后不捕获也不注入。 |
| `capture_files` | bool | `True` | 开启后会把 docx/xlsx/pptx/pdf/txt/md/csv 等文件的正文抽出来入库，跨会话可检索。 |
| `capture_tools` | bool | `True` | 记录每次工具调用的名称、参数与结果摘要，方便日后追溯'你之前处理过什么'。 |
| `inject_enabled` | bool | `True` | 每轮对话按当前消息检索全局记忆，命中则作为临时上下文注入，不写入历史、不干扰人格。 |
| `inject_top_k` | int | `3` | 建议 2-4 条，太多会挤占上下文。 |
| `inject_max_chars` | int | `1200` | 控制注入块总长度，防止长文档挤爆上下文。 |
| `file_text_max_chars` | int | `6000` | 超过会被截断。检索时只注入命中片段，不受此限制影响存储。 |
| `max_file_mb` | int | `30` | 超过该大小的文件只记录元信息，不解析正文。 |

## 依赖

无第三方依赖（纯标准库）

## 兼容性

- 平台：*
- AstrBot：`>=4.24,<5`

## 隐私说明

插件只在本机运行，不上传任何数据到第三方。配置与运行时数据存在 `AstrBot/data/` 下。

## 许可

MIT，见 [LICENSE](LICENSE)。
