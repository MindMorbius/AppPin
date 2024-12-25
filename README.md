# AppPin

利用图像识别、OCR框架，识别Windows的软件界面，并处理信息，通过模拟操作（复制、粘贴）按照用户需求控制软件。

- 虚拟桌面方案
- openCV

## 微信
针对微信（Windows）软件，需要实现联系人消息的收发。

1. 转发消息
Bot识别到微信新消息，通过模拟点击进入聊天窗口，通过右键—复制将消息转发到IM软件中。

2. 接收消息
用户在IM软件中的回复消息，Bot在微信软件中点开对应联系人的聊天窗口，通过右键—粘贴，将消息发送到文本框中，点击发送按钮。


## 使用

1. 创建虚拟环境
```
python -m venv venv
```

2. 激活虚拟环境
```
venv\Scripts\activate
```

3. 安装依赖
```
pip install -r requirements.txt
```

4. 运行
```
python wechat_monitor.py
```
