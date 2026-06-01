# Photoshop scripts

## 如何使用 Photoshop scripts：

1. 下载脚本，保存到任意位置；
2. 在 Photoshop 中，选择 File -> Scripts -> Browse... ，然后选中对应脚本；  
3. 点击回车，运行脚本。

## psd_link__batchConverToLinked.jsx

批量将当前 psd 文件中内嵌的智能对象导出到当前文件所在目录的 ./_LinkedObjects 目录中。

当前脚本默认自动去重， 不会导出重复的智能对象 ；去重方式有两种， 一种是根据同一个智能对象去重，另外一个是根据相同内容进行去重。