#target photoshop

// 1. 定义符号/文字与对齐方式的字典映射
var alignDict = {
    "^": "top", "v": "bottom", "<": "left", ">": "right", "*": "middle",
    "top": "top", "bottom": "bottom", "left": "left", "right": "right", "middle": "middle", "center": "middle",
    "上": "top", "下": "bottom", "左": "left", "右": "right", "中": "middle"
};

var opDict = {
    "stretch": "stretch", "fit": "fit", "fill": "fill",
    "拉伸": "stretch", "变形": "stretch",
    "适应": "fit", "填充": "fill"
};

// 全局配置（由运行初期的 Dialog 进行统一初始化）
var globalConfig = {
    bitmapAction: "convert", // "convert" (转智能对象), "direct" (直接拉伸), "skip" (跳过)
    transformMode: "stretch"  // 缺省全局变换模式："stretch", "fit", "fill"
};

function main() {
    if (app.documents.length === 0) {
        alert("请先打开一个 Photoshop 文档！");
        return;
    }

    var doc = app.activeDocument;
    
    // 步骤一：预扫描文档，检查是否存在两端对齐以及是否包含 Bitmap/智能对象
    var scanResult = preScanDoc(doc);
    
    // 如果存在需要两端对齐的图层，弹出 Dialog 进行统一策略配置
    if (scanResult.hasAlignBoth && !scanResult.allAlignBothHaveValidSuffixOp) {
        if (!showConfigDialog(scanResult.hasBitmapOrSmart)) {
            return; // 用户点击取消
        }
    }

    // 步骤二：执行自底向上的递归对齐与变换
    doc.suspendHistory("自底向上边界对齐变换", "processContainerBottomUp(doc)");
    alert("所有的边界自底向上对齐与变换处理完成！");
}

// 预扫描：计算是否需要弹出选项
function preScanDoc(container) {
    var res = { hasAlignBoth: false, hasBitmapOrSmart: false, allAlignBothHaveValidSuffixOp: true };
    var regex = /.*\[\[([^\]]+)\]\](?::([a-zA-Z\u4e00-\u9fa5]+))?.*/;
    
    function scan(obj) {
        var layers = obj.layers;
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            if (layer.typename === "LayerSet") {
                scan(layer);
            } else {
                if (regex.test(layer.name)) {
                    var match = layer.name.match(regex);
                    var cmdStr = match[1].toLowerCase();
                    var suffixOp = match[2] ? match[2].toLowerCase() : null;
                    var alignments = parseAlignments(cmdStr);
                    if ((alignments.left && alignments.right) || (alignments.top && alignments.bottom)) {
                        res.hasAlignBoth = true;
                        if (!(suffixOp && opDict[suffixOp])) res.allAlignBothHaveValidSuffixOp = false;
                        var parentGroup = layer.parent;
                        for (var j = 0; j < parentGroup.layers.length; j++) {
                            var sibling = parentGroup.layers[j];
                            if (sibling !== layer) {
                                var t = getLayerType(sibling);
                                if (t === "bitmap" || t === "smart") res.hasBitmapOrSmart = true;
                            }
                        }
                    }
                }
            }
        }
    }
    scan(container);
    return res;
}

// 弹出 UI 配置面板
function showConfigDialog(showBitmapOptions) {
    var dlg = new Window("dialog", "边界对齐与变换统一操作配置");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    
    // 变换模式面板
    var pMode = dlg.add("panel", undefined, "未指明后缀时的缺省两端对齐变换模式");
    pMode.orientation = "row";
    pMode.padding = 15;
    var rbStretch = pMode.add("radiobutton", undefined, "拉伸变形 (Stretch)");
    var rbFit = pMode.add("radiobutton", undefined, "等比适应 (Fit)");
    var rbFill = pMode.add("radiobutton", undefined, "等比填充 (Fill)");
    rbStretch.value = true;
    
    // Bitmap 处理面板
    var pBitmap, rbConvert, rbDirect, rbSkip;
    if (showBitmapOptions) {
        pBitmap = dlg.add("panel", undefined, "针对像素图层(Bitmap)及尺寸不足的智能对象的处理方案");
        pBitmap.orientation = "column";
        pBitmap.alignChildren = "left";
        pBitmap.padding = 15;
        rbConvert = pBitmap.add("radiobutton", undefined, "转换为智能对象保护画质 (推荐)");
        rbDirect = pBitmap.add("radiobutton", undefined, "不转换，允许直接对其像素拉伸变换");
        rbSkip = pBitmap.add("radiobutton", undefined, "跳过变换操作 (仅对它们做单向或居中对齐)");
        rbConvert.value = true;
    }
    
    var pBtns = dlg.add("group");
    pBtns.alignment = "center";
    var btnOk = pBtns.add("button", undefined, "执行扫描对齐", {name: "ok"});
    var btnCancel = pBtns.add("button", undefined, "取消", {name: "cancel"});
    
    if (dlg.show() == 1) {
        if (rbStretch.value) globalConfig.transformMode = "stretch";
        if (rbFit.value) globalConfig.transformMode = "fit";
        if (rbFill.value) globalConfig.transformMode = "fill";
        
        if (showBitmapOptions) {
            if (rbConvert.value) globalConfig.bitmapAction = "convert";
            if (rbDirect.value) globalConfig.bitmapAction = "direct";
            if (rbSkip.value) globalConfig.bitmapAction = "skip";
        }
        return true;
    }
    return false;
}

// 核心逻辑：自底向上（从最里层向外）处理
function processContainerBottomUp(container) {
    var layers = container.layers;
    
    // 1. 先完全深度递归进入内部子组
    for (var i = layers.length - 1; i >= 0; i--) {
        var layer = layers[i];
        if (layer.typename === "LayerSet") {
            processContainerBottomUp(layer); 
        }
    }
    
    // 2. 内部处理完毕后，再检查当前层级是否有 Bound 边界层
    var boundLayer = findBoundLayerDirect(container);
    if (boundLayer) {
        alignGroupContent(container, boundLayer);
    }
}

// 仅在当前层级查找 Bound 层（不向下深搜，防作用域错乱）
function findBoundLayerDirect(group) {
    var subLayers = group.layers;
    var regex = /.*\[\[([^\]]+)\]\](?::([a-zA-Z\u4e00-\u9fa5]+))?.*/;
    for (var i = 0; i < subLayers.length; i++) {
        if (subLayers[i].typename !== "LayerSet" && regex.test(subLayers[i].name)) {
            return subLayers[i];
        }
    }
    return null;
}

function parseAlignments(cmdStr) {
    var alignments = { top: false, bottom: false, left: false, right: false, middle: false };
    for (var key in alignDict) {
        if (alignDict.hasOwnProperty(key) && cmdStr.indexOf(key.toLowerCase()) !== -1) {
            alignments[alignDict[key]] = true;
        }
    }
    return alignments;
}

// 精准识别图层类型
function getLayerType(layer) {
    if (layer.typename === "LayerSet") return "group";
    if (layer.kind === LayerKind.TEXT) return "text";
    
    // 通过 ActionManager 侦测智能对象
    try {
        var ref = new ActionReference();
        ref.putIndex(charIDToTypeID('Lyr '), layer.itemIndex);
        var desc = executeActionGet(ref);
        if (desc.hasKey(stringIDToTypeID('smartObject'))) {
            return "smart";
        }
    } catch(e) {}
    
    if (layer.kind === LayerKind.NORMAL) return "bitmap";
    return "shape"; // 文本、纯色、渐变或矢量图层归类为直接变换层
}

// 获取智能对象文件的本体宽高分辨率，检测够不够变
function isSmartObjectLargeEnough(layer, targetW, targetH) {
    try {
        var ref = new ActionReference();
        ref.putIndex(charIDToTypeID('Lyr '), layer.itemIndex);
        var desc = executeActionGet(ref);
        if (desc.hasKey(stringIDToTypeID('smartObject'))) {
            var smDesc = desc.getObjectValue(stringIDToTypeID('smartObject'));
            var sizeDesc = smDesc.getObjectValue(stringIDToTypeID('size'));
            var originalW = sizeDesc.getDouble(stringIDToTypeID('width'));
            var originalH = sizeDesc.getDouble(stringIDToTypeID('height'));
            
            // 如果内部素材本体大于或等于目标所需大小，说明无损放大储备足够
            if (originalW >= targetW && originalH >= targetH) {
                return true;
            }
        }
    } catch(e) {}
    return false;
}

// 将图层封包转为智能对象
function convertToSmartObject(layer) {
    var doc = app.activeDocument;
    var active = doc.activeLayer;
    doc.activeLayer = layer;
    try {
        var idnewPlacedLayer = stringIDToTypeID("newPlacedLayer");
        executeAction(idnewPlacedLayer, undefined, DialogModes.NO);
    } catch(e) {}
    doc.activeLayer = active;
}

// 对目标组执行具体的对齐和拉伸变换变换
function alignGroupContent(group, boundLayer) {
    var regex = /.*\[\[([^\]]+)\]\](?::([a-zA-Z\u4e00-\u9fa5]+))?.*/;
    var match = boundLayer.name.match(regex);
    if (!match) return;
    
    var cmdStr = match[1].toLowerCase();
    var suffixOp = match[2] ? match[2].toLowerCase() : null;
    
    // 决定采用何种变换操作模式（优先解析后缀标签，否则使用全局缺省）
    var currentTransformMode = globalConfig.transformMode;
    if (suffixOp && opDict[suffixOp]) {
        currentTransformMode = opDict[suffixOp];
    }
    
    var alignments = parseAlignments(cmdStr);
    
    // 获取 Boundary 矩形属性
    var bBounds = boundLayer.bounds;
    var bL = parseFloat(bBounds[0]);
    var bT = parseFloat(bBounds[1]);
    var bR = parseFloat(bBounds[2]);
    var bB = parseFloat(bBounds[3]);
    var bW = bR - bL;
    var bH = bB - bT;
    
    var subLayers = group.layers;
    
    // 逆序遍历组内元素进行定位和变换
    for (var i = subLayers.length - 1; i >= 0; i--) {
        var targetLayer = subLayers[i];
        if (targetLayer === boundLayer) continue;
        
        var tBounds = targetLayer.bounds;
        var tL = parseFloat(tBounds[0]);
        var tT = parseFloat(tBounds[1]);
        var tR = parseFloat(tBounds[2]);
        var tB = parseFloat(tBounds[3]);
        var tW = tR - tL;
        var tH = tB - tT;
        
        var deltaX = 0;
        var deltaY = 0;
        
        var isAlignBothX = (alignments.left && alignments.right);
        var isAlignBothY = (alignments.top && alignments.bottom);
        
        // --- 1. 基础位移计算 (计算图层对齐所需偏移) ---
        if (isAlignBothX || alignments.middle) {
            deltaX = (bL + bW / 2) - (tL + tW / 2);
        } else if (alignments.left) {
            deltaX = bL - tL;
        } else if (alignments.right) {
            deltaX = bR - tR;
        }
        
        if (isAlignBothY || alignments.middle) {
            deltaY = (bT + bH / 2) - (tT + tH / 2);
        } else if (alignments.top) {
            deltaY = bT - tT;
        } else if (alignments.bottom) {
            deltaY = bB - tB;
        }
        
        if (deltaX !== 0 || deltaY !== 0) {
            targetLayer.translate(deltaX, deltaY);
            // 同步修正当前图层的临时坐标值
            tL += deltaX; tR += deltaX; tT += deltaY; tB += deltaY;
        }
        
        // --- 2. 两端对齐的高级缩放变换逻辑 ---
        if (isAlignBothX || isAlignBothY) {
            var lType = getLayerType(targetLayer);
            
            // 安全机制：检验或拦截非矢量/Bitmap类图层
            if (lType === "bitmap") {
                if (globalConfig.bitmapAction === "skip") continue; 
                if (globalConfig.bitmapAction === "convert") {
                    convertToSmartObject(targetLayer);
                    targetLayer = group.layers[i]; // 重新装载包装后的智能对象引用
                }
            } else if (lType === "smart") {
                // 如果是智能对象，深入探测其原片尺寸是否可以胜任本次变换
                var isEnough = isSmartObjectLargeEnough(targetLayer, bW, bH);
                if (!isEnough && globalConfig.bitmapAction === "skip") {
                    continue; // 原始尺寸不足且用户要求跳过
                }
            }
            
            var scaleX = 100;
            var scaleY = 100;
            
            // 根据三种变换模型分别计算缩放比
            if (currentTransformMode === "stretch") {
                if (isAlignBothX) scaleX = (bW / tW) * 100;
                if (isAlignBothY) scaleY = (bH / tH) * 100;
            } else if (currentTransformMode === "fit") {
                var ratioX = isAlignBothX ? (bW / tW) : (bH / tH);
                var ratioY = isAlignBothY ? (bH / tH) : (bW / tW);
                var minRatio = Math.min(bW / tW, bH / tH);
                scaleX = minRatio * 100;
                scaleY = minRatio * 100;
            } else if (currentTransformMode === "fill") {
                var maxRatio = Math.max(bW / tW, bH / tH);
                scaleX = maxRatio * 100;
                scaleY = maxRatio * 100;
            }
            
            // 执行中心锚点无损/有损缩放
            if (scaleX !== 100 || scaleY !== 100) {
                targetLayer.resize(scaleX, scaleY, AnchorPosition.MIDDLECENTER);
            }
        }
    }
}

main();
