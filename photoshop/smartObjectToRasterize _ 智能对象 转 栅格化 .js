// 1. 强制设定对话框模式为无
app.displayDialogs = DialogModes.NO;

/**
 * 核心逻辑函数：必须放在最外层（全局），确保 suspendHistory 能找到它
 */
function processAllSmartObjects() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;

    // 定义递归逻辑
    var recursiveRasterize = function (layers) {
        for (var i = layers.length - 1; i >= 0; i--) {
            var layer = layers[i];

            if (layer.typename === "LayerSet") {
                recursiveRasterize(layer.layers);
            } else if (layer.kind === LayerKind.SMARTOBJECT) {
                // 记录原始可见性
                var originalVisibility = layer.visible;

                try {
                    // 执行栅格化
                    layer.rasterize(RasterizeType.ENTIRELAYER);

                    // 还原可见性
                    if (layer.visible !== originalVisibility) {
                        layer.visible = originalVisibility;
                    }
                } catch (e) {
                    $.writeln("跳过图层: " + layer.name);
                }
            }
        }
    };

    recursiveRasterize(doc.layers);
}

/**
 * 入口函数
 */
function main() {
    if (app.documents.length === 0) {
        alert("请先打开文档");
        return;
    }

    var doc = app.activeDocument;

    // 关键修正点：
    // 1. 确保函数名是全局可访问的
    // 2. 使用字符串格式并加上括号 "functionName()"
    try {
        doc.suspendHistory("批量栅格化(保持隐藏状态)", "processAllSmartObjects()");
        alert("处理完成！所有智能对象已转为位图并保留原可见性。");
    } catch (err) {
        alert("执行失败: " + err);
    }
}

// 启动
main();