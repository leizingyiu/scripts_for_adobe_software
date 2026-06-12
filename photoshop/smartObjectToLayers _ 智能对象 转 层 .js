// 开启快速处理模式
app.displayDialogs = DialogModes.NO;

// 1. 定义一个专门的全局执行函数供 suspendHistory 调用
function runBatchConvertProcess() {
    // 在这个函数内部重新获取 activeDocument，确保作用域安全
    if (app.documents.length > 0) {
        processLayers(app.activeDocument.layers);
    }
}

/**
 * 递归遍历图层函数
 */
function processLayers(layers) {
    for (var i = layers.length - 1; i >= 0; i--) {
        var currentLayer = layers[i];

        if (currentLayer.typename === "LayerSet") {
            processLayers(currentLayer.layers);
        }
        else if (currentLayer.kind === LayerKind.SMARTOBJECT) {
            // 记录原始可见性
            var wasVisible = currentLayer.visible;

            app.activeDocument.activeLayer = currentLayer;

            try {
                var idplacedLayerConvertToLayers = stringIDToTypeID("placedLayerConvertToLayers");
                executeAction(idplacedLayerConvertToLayers, undefined, DialogModes.NO);

                // 还原可见性：转换后的新图层/组会自动选中，直接操作 activeLayer
                if (app.activeDocument.activeLayer.visible !== wasVisible) {
                    app.activeDocument.activeLayer.visible = wasVisible;
                }
            } catch (e) {
                $.writeln("无法转换图层: " + currentLayer.name);
            }
        }
    }
}

function main() {
    if (app.documents.length === 0) {
        alert("请先打开一个文档。");
        return;
    }

    // 2. 关键修正：传递全局函数的名称字符串，并加上括号
    // 这样 PS 的历史记录引擎就能正确回调它
    app.activeDocument.suspendHistory("批量所有智能对象转图层", "runBatchConvertProcess()");

    alert("处理完成。");
}

// 执行脚本
main();