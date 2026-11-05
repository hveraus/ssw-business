# SSW 商务

SSW（智鹏瑞尔软件）企业 AI 转型业务材料。

> **要改这份 deck？先读 [AGENTS.md](AGENTS.md)。** 里面有页码机制、三个已知的坑、设计系统、内容口径铁律，以及"改一页/加一页/删一页"的具体做法。

## 内容

| 路径 | 说明 |
|---|---|
| `SSW AI 业务_v2/SSW AI 业务_v7.pptx` | **当前版本**（62 页）对外提案，由脚本生成 |
| `SSW AI 业务_v2/build_deck.py` | deck 的**唯一源头**——所有改动走脚本 |
| `SSW AI 业务_v2/check_deck.py` | 产出校验器（页码 / 版面 / 换行编码），提交前必须跑 |
| `SSW AI 业务_v2/assets/` | 图片资产（世界地图、冰山、界面截图） |
| `SSW AI 业务_优化大纲_v2.md` | 内容大纲 |

## 构建与校验

```bash
cd "SSW AI 业务_v2"
python3 build_deck.py     # 生成 SSW AI 业务_v7.pptx
python3 check_deck.py     # 必须全绿（退出码 0）才能提交
```

依赖 `python-pptx`。

## 注意

- **不要手工改 pptx 里的文字**——下次重建会被覆盖。改 `build_deck.py`。
- 本仓库为私有：含客户名称（已脱敏）、代采折扣与计价口径，**请勿转为公开**。
