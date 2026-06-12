// 开启快速处理模式
app.displayDialogs = DialogModes.NO;

/**
 * 核心逻辑函数：必须放在全局作用域，以便 suspendHistory 调用
 */
function runSmartObjectToArtboardProcess() {
    var doc = app.activeDocument;
    var layers = doc.layers;
    var smartObjects = [];

    // 1. 收集所有的智能对象（从后往前收集，防止操作后索引混乱）
    for (var i = layers.length - 1; i >= 0; i--) {
        if (layers[i].kind == LayerKind.SMARTOBJECT) {
            smartObjects.push(layers[i]);
        }
    }

    if (smartObjects.length === 0) {
        return;
    }

    // 2. 遍历并处理
    for (var j = 0; j < smartObjects.length; j++) {
        var targetLayer = smartObjects[j];

        // 记录原始可见性
        var wasVisible = targetLayer.visible;

        doc.activeLayer = targetLayer;

        // 获取图层边界
        var bounds = targetLayer.bounds;

        // 创建画板
        makeArtboardFromLayer(targetLayer.name, bounds);

        // 3. 执行 PS 内置命令：转换为图层
        try {
            executeAction(stringIDToTypeID("placedLayerConvertToLayers"), undefined, DialogModes.NO);

            // 转换后，如果原智能对象是隐藏的，我们需要把新生成的图层组/图层也隐藏
            // 因为转换命令通常会使结果变为可见
            if (!wasVisible) {
                doc.activeLayer.visible = false;
            }
        } catch (e) {
            continue;
        }
    }
}

/**
 * 主入口
 */
function main() {
    if (app.documents.length === 0) {
        alert("请先打开一个包含智能对象的文档。");
        return;
    }

    var doc = app.activeDocument;

    // 关键修正：传入全局函数名的字符串
    doc.suspendHistory("批量智能对象转画板", "runSmartObjectToArtboardProcess()");

    alert("处理完成。");
}

/**
 * Action Manager API 创建画板
 */
function makeArtboardFromLayer(name, bounds) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass(stringIDToTypeID("artboardSection"));
    desc1.putReference(stringIDToTypeID("null"), ref1);

    var desc2 = new ActionDescriptor();
    desc2.putString(stringIDToTypeID("name"), name);
    desc1.putObject(stringIDToTypeID("using"), stringIDToTypeID("artboardSection"), desc2);

    var desc3 = new ActionDescriptor();
    desc3.putDouble(stringIDToTypeID("top"), bounds[1].value);
    desc3.putDouble(stringIDToTypeID("left"), bounds[0].value);
    desc3.putDouble(stringIDToTypeID("bottom"), bounds[3].value);
    desc3.putDouble(stringIDToTypeID("right"), bounds[2].value);
    desc1.putObject(stringIDToTypeID("artboardRect"), stringIDToTypeID("classFloatRect"), desc3);

    executeAction(stringIDToTypeID("make"), desc1, DialogModes.NO);
}

main();