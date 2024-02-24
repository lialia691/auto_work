这是虚拟环境建立的
迁移的话好迁移

你可以使用pip freeze > requirements.txt命令来将你的依赖包和对应的版本储存到一个文本文件中，这个文件常常被命名为requirements.txt。
然后，你就可以将这个requirements.txt一并提交到git仓库。当其他人clone你的代码并且使用pip install -r requirements.txt命令时，pip会自动安装requirements.txt中记载的所有依赖包和对应的版本。

16:38 2024/2/23
已经添加到git
22:38 2024/2/23
关于编码 
git中修改编码来显示中文:
我们需要设置让Git使用UTF-8字符编码：
首先，找到Git安装路径下的etc文件夹，并在其下找到gitconfig文件，这个就是Git的全局配置文件。
使用记事本或者其他的文本编辑器打开它。
在[core]部分下加入quotepath = false设置。

若cmd终端不显示中文,需这样设置:
以Windows操作系统下的cmd命令行为例，可以使用chcp 65001命令来设置控制台字符编码为UTF-8，这样就可以在控制台中正常显示中文字符了。操作步骤如下：
打开cmd命令行。
输入chcp 65001，然后按回车。
23:15 2024/2/23
