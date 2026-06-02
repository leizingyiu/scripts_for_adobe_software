# Photoshop scripts

## 如何使用 Photoshop scripts：

1. 下载脚本，保存到任意位置；
2. 在 Photoshop 中，选择 File -> Scripts -> Browse... ，然后选中对应脚本；  
3. 点击回车，运行脚本。

## psd_link__batchConverToLinked.jsx

批量将当前 psd 文件中内嵌的智能对象导出到当前文件所在目录的 ./_LinkedObjects 目录中。

当前脚本默认自动去重， 不会导出重复的智能对象 ；去重方式有两种， 一种是根据同一个智能对象去重，另外一个是根据相同内容进行去重。

## ps_status.jsx

用于在 PSD 顶层 `status` 组内批量生成/更新不同 UI 状态。状态层名里写 `{K:V,...}` 配置，脚本会扫描所有以冒号结尾的 Key 组/文本层（支持 select/toggle/textContent 等结构），按 value 切换显隐或写入文本，并把最终画面回写到状态层。支持一键更新全部/选中状态层，也可从 Excel 复制 TSV 批量创建并刷新，适合按钮、组件多状态的快速出图。
