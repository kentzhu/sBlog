关于sBlog
--
高中时候写的一个ASP博客程序。
简简单单。
作为纪念吧。

使用
--

默认用户名密码
-------

户名：admin
密码：admin888

安装
-------
sBlog的安装很简单，只要将压缩包解压到您需要安装的目录下即可。不过您需要做一件事，就是修改数据库名称。

首先，进入Data文件夹，将sBlog_Basic.mdb文件改成任意文件名，如oiasl _scB g B.mdb。这样做是为了防止他人通过默认数据库名称下载到您的数据库。
然后，打开根目录的comm.asp文件，找到第三行

    StrDbName="Data/sBlog_Basic.mdb"        '数据库名称以及路径

将其中的文件名sBlog_Basic.mdb更换为您修改后的文件名。如：

    StrDbName="Data/oiasl_scB#g#B.mdb"        '数据库名称以及路径

修改完毕后，打开浏览器，进入Blog。

接下来用默认的用户名和密码进入后台，在第一时间修改您的密码。然后点击侧边栏的“基础设置”修改博客的各项设置信息。

