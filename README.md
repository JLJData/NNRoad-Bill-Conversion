# NNRoad 账单转换服务（Python）

Office 通过 HTTP 调用本仓；各引擎逻辑在 profiles/，公共 mapping 在 ill_convert/ 与 convert_mapping.py。

## 启动

Linux / 服务器：`pip install -r requirements.txt`

本机 Windows（需要 Excel COM 快照时）：`pip install -r requirements-windows.txt`

`启动转换服务.bat`

或 `python -m uvicorn convert_api:app --host 127.0.0.1 --port 8765`

## 主要目录

- convert_api.py — HTTP 接口
- convert_runner.py / engines.py — 引擎调度
- convert_mapping.py / mapping_inspect.py — Office 映射
- ill_convert/ — 多引擎公共库
- profiles/ — 各地区转换实现
- 	emplates/ — 地区默认 PN 母版

本地样例与试转输出放 账单/、输出/（已 gitignore）。
