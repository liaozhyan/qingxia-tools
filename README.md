# 轻匣小工具（Qingxia Tools）

轻匣（https://www.iqingxia.com/）出品的 Windows 会务/办公小工具集合源码。

## 包含的工具

- 会务助手（主程序）
- 会务视频、会务音乐
- PPT 转换、PPT 倒计时
- 座位牌打印、现场抽签排序
- 扫码收文件、长条屏合成、二维码生成
- 提词器、局域网互传、图片工具箱、音视频剪辑、截图

入口脚本统一放在 `小工具/` 目录，公共模块放在仓库根目录。

## 运行环境

- Windows 10/11
- Python 3.11 及以上（建议 3.12）
- 安装依赖：

```bash
pip install -r requirements.txt
```

## 运行

```bash
python 小工具/会务音乐.py
python 小工具/会务视频.py
python 小工具/PPT转换.py
```

其他工具的入口脚本在 `小工具/` 下，文件名即工具名。

## 打包

```bash
pip install pyinstaller
pyinstaller -F -w 小工具/PPT转换.py
```

## 许可证

本项目采用 MIT 许可证，详见 [LICENSE](LICENSE)。

第三方依赖（Pillow、NumPy、OpenCV、imageio-ffmpeg 等）遵循其各自的开源许可证。
