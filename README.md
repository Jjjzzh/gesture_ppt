# Gesture PPT Controller

基于 OpenCV + MediaPipe 的手势控制 PPT 演示系统。

## 效果展示

- 摄像头实时捕捉手部 21 个关键点，霓虹风格 HUD 界面
- 手势识别状态实时显示 + 动作触发全屏闪烁动画

## 手势操作

| 手势 | 动作 | 说明 |
|------|------|------|
| 五指张开 ✋ | 开始放映 | 手掌完全展开 |
| 比 1 ☝️ | 下一页 | 仅食指伸出 |
| 比 2 ✌️ | 上一页 | 食指+中指伸出 |

按 Q 或 ESC 退出。

## 环境要求

- Python 3.10+
- Microsoft PowerPoint
- 摄像头

## 快速开始

```bash
# 1. 安装依赖
pip install opencv-python mediapipe pywin32 numpy

# 2. 打开你的 PPT 文件（PowerPoint 需保持运行）

# 3. 运行手势控制
python gesture_ppt.py
```

首次运行会自动下载手部关键点模型（约 7.8MB），缓存于 `C:/temp/`。

## 与 PPT 的集成方式

通过 Windows COM 接口直接控制 PowerPoint，**不依赖键盘模拟**，无需担心窗口焦点问题：

- 脚本自动连接到已运行的 PowerPoint 实例
- `SlideShowSettings.Run()` → 启动全屏放映
- `View.Next()` / `View.Previous()` → 翻页
- 即使摄像头窗口在前台，也能正常控制 PPT

## 使用步骤

1. 打开 PowerPoint 文件（确保 PowerPoint 在运行）
2. 运行 `python gesture_ppt.py` → 控制台显示 `[PPT] Connected to PowerPoint`
3. **握拳保持 0.3 秒** → 自动进入全屏放映
4. **大拇指朝上** → 下一页；**大拇指朝下** → 上一页
5. 按 Q 退出

## 技术栈

- **OpenCV** — 摄像头采集与图像渲染
- **MediaPipe Tasks** — 手部 21 点关键点实时检测
- **pywin32 (COM)** — 直接操控 PowerPoint 对象模型
- **NumPy** — 矩阵运算与动画计算

## 防误触机制

- 1.2 秒冷却间隔，避免连续触发
- 手势需保持连续 N 帧确认（握拳 8 帧，大拇指 5 帧）
- 大拇指手势需 4 个手指弯曲 + 拇指明显伸出 + 方向明确
