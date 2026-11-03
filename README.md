# SSW 商务

SSW（智鹏瑞尔软件）企业 AI 转型业务材料。

## 内容

| 路径 | 说明 |
|---|---|
| `SSW AI 业务_v2/SSW AI 业务_v3.pptx` | **当前版本**（57 页）对外提案 |
| `SSW AI 业务_v2/SSW AI 业务_v2.pptx` | 上一版，留档 |
| `SSW AI 业务_v2/build_deck.py` | deck 的**唯一源头**——所有文案改动走脚本，改完重新生成 |
| `SSW AI 业务_v2/assets/` | 图片资产（世界地图、冰山、界面截图） |
| `SSW AI 业务_优化大纲_v2.md` | 内容大纲 |

## 重新生成 deck

```bash
cd "SSW AI 业务_v2"
python3 build_deck.py
```

依赖 `python-pptx`。

## 注意

- **不要手工改 pptx 里的文字**——下次重建会被覆盖。改 `build_deck.py`。
- 本仓库为私有：含客户名称（已脱敏）、代采折扣与加价口径，请勿转为公开。
