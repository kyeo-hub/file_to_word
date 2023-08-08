# 利用固定模板文件生成Word

> 使用tkinter作为GUI的python程序，类似邮件合并，数据源没有使用Excel，而是使用了SeaTable的API（当然换成Notion等也应该可以，本例没有使用）

- file_to_word.py 主程序
- export_word.ico 程序中使用的图标
- m.docx 制作好的模板文件
- requirements.txt pip依赖包管理
## 安装依赖包
```
pip install -r requirements.txt
```
## 运行程序
```
python file_to_word.py
```
#### 程序截图、Seatable截图、Word模板截图
|Seatable内容|程序主界面|Word邮件合并模板|
|--|--|--|
|![SeaTable](https://image.10an.fun/%E4%BC%81%E4%B8%9A%E5%BE%AE%E4%BF%A1%E6%88%AA%E5%9B%BE_16914753847794.png)|![程序主界面](https://image.10an.fun/%E4%BC%81%E4%B8%9A%E5%BE%AE%E4%BF%A1%E6%88%AA%E5%9B%BE_16914755487356.png)|![Word模板](https://image.10an.fun/%E4%BC%81%E4%B8%9A%E5%BE%AE%E4%BF%A1%E6%88%AA%E5%9B%BE_16914756802232.png)|
## 程序打包
 ```
 pip install pyinstaller
 pyinstaller -F -w file_to_word.py
 ```

```
Pyinstaller常用参数 含义
-i 或 -icon 生成icon
-F 创建一个绑定的可执行文件
-w 使用窗口，无控制台
-C 使用控制台，无窗口
-D 创建一个包含可执行文件的单文件夹包(默认情况下)
-n 文件名
```

