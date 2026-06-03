#target photoshop

function main() {
    // 检查是否有打开的文档
    if (app.documents.length === 0) {
        alert("请先打开一个 Photoshop 文档！");
        return;
    }

    var doc = app.activeDocument;

    // --- UI 界面构建 ---
    var win = new Window("dialog", "图层批量重命名工具");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 10;
    win.margins = 16;

    // 1. 操作类型下拉菜单
    var groupType = win.add("group");
    groupType.add("statictext", undefined, "选择操作:");
    var dropDown = groupType.add("dropdownlist", undefined, ["在前面添加", "在后面添加", "查找替换", "正则查找替换"]);
    dropDown.selection = 0;
    dropDown.size = [150, 25];

    // 2. 动态输入框面板
    var panelInput = win.add("panel", undefined, "参数设置");
    panelInput.orientation = "column";
    panelInput.alignChildren = ["fill", "top"];
    panelInput.spacing = 8;

    // 创建所有可能用到的输入组件，后面根据下拉菜单显示/隐藏
    var groupIn1 = panelInput.add("group");
    var lblIn1 = groupIn1.add("statictext", undefined, "文本内容:");
    var txtIn1 = groupIn1.add("edittext", undefined, "");
    txtIn1.alignment = ["fill", "center"];

    var groupIn2 = panelInput.add("group");
    var lblIn2 = groupIn2.add("statictext", undefined, "替换为:  ");
    var txtIn2 = groupIn2.add("edittext", undefined, "");
    txtIn2.alignment = ["fill", "center"];

    // 3. 结果预演区域
    win.add("statictext", undefined, "检查/预演结果:");
    var txtResult = win.add("edittext", undefined, "", {multiline: true, scrolling: true});
    txtResult.size = [400, 150];

    // 4. 底部按钮区域
    var groupButtons = win.add("group");
    groupButtons.alignment = "right";
    var btnCheck = groupButtons.add("button", undefined, "检查 (预演)");
    var btnRename = groupButtons.add("button", undefined, "重命名图层名", {name: "ok"});
    var btnCancel = groupButtons.add("button", undefined, "取消", {name: "cancel"});

    // --- UI 逻辑切换 ---
    function updateUI() {
        if (dropDown.selection.index === 0 || dropDown.selection.index === 1) {
            // 前面添加 或 后面添加
            lblIn1.text = "添加文本:";
            groupIn2.visible = false;
        } else {
            // 查找替换 或 正则查找替换
            lblIn1.text = "查找内容:";
            groupIn2.visible = true;
        }
        win.layout.layout(true); // 重新布局窗口
    }
    
    dropDown.onChange = updateUI;
    updateUI(); // 初始化UI状态

    // --- 核心业务逻辑 ---

    // 递归获取所有图层（当没有选中图层，处理全部图层时使用）
    function getAllLayers(layerContainer) {
        var layers = [];
        for (var i = 0; i < layerContainer.layers.length; i++) {
            var layer = layerContainer.layers[i];
            layers.push(layer);
            if (layer.typename === "LayerSet") {
                layers = layers.concat(getAllLayers(layer));
            }
        }
        return layers;
    }

    function getSelectedLayerIDs() {
        var ids = [];

        function addId(id) {
            for (var i = 0; i < ids.length; i++) {
                if (ids[i] === id) return;
            }
            ids.push(id);
        }

        try {
            var refIDs = new ActionReference();
            refIDs.putProperty(stringIDToTypeID("property"), stringIDToTypeID("targetLayersIDs"));
            refIDs.putEnumerated(stringIDToTypeID("document"), stringIDToTypeID("ordinal"), stringIDToTypeID("targetEnum"));
            var descIDs = executeActionGet(refIDs);

            if (descIDs.hasKey(stringIDToTypeID("targetLayersIDs"))) {
                var listIDs = descIDs.getList(stringIDToTypeID("targetLayersIDs"));
                for (var i = 0; i < listIDs.count; i++) {
                    try {
                        addId(listIDs.getInteger(i));
                    } catch (e) {
                        addId(listIDs.getReference(i).getIdentifier());
                    }
                }
            }
        } catch (e) {}

        try {
            var refIDs2 = new ActionReference();
            refIDs2.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("targetLayersIDs"));
            refIDs2.putEnumerated(charIDToTypeID("Dcmn"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
            var descIDs2 = executeActionGet(refIDs2);

            if (descIDs2.hasKey(stringIDToTypeID("targetLayersIDs"))) {
                var listIDs2 = descIDs2.getList(stringIDToTypeID("targetLayersIDs"));
                for (var i = 0; i < listIDs2.count; i++) {
                    try {
                        addId(listIDs2.getInteger(i));
                    } catch (e) {
                        addId(listIDs2.getReference(i).getIdentifier());
                    }
                }
            }
        } catch (e) {}

        try {
            var refTargets = new ActionReference();
            refTargets.putProperty(stringIDToTypeID("property"), stringIDToTypeID("targetLayers"));
            refTargets.putEnumerated(stringIDToTypeID("document"), stringIDToTypeID("ordinal"), stringIDToTypeID("targetEnum"));
            var descTargets = executeActionGet(refTargets);

            if (descTargets.hasKey(stringIDToTypeID("targetLayers"))) {
                var listTargets = descTargets.getList(stringIDToTypeID("targetLayers"));
                for (var i = 0; i < listTargets.count; i++) {
                    var refItem = listTargets.getReference(i);
                    try {
                        addId(refItem.getIdentifier());
                    } catch (e) {
                        var itemIndex = refItem.getIndex();
                        var refByIndex = new ActionReference();
                        refByIndex.putIndex(charIDToTypeID("Lyr "), itemIndex);
                        var descByIndex = executeActionGet(refByIndex);
                        if (descByIndex.hasKey(stringIDToTypeID("layerID"))) {
                            addId(descByIndex.getInteger(stringIDToTypeID("layerID")));
                        }
                    }
                }
            }
        } catch (e) {}

        try {
            var refTargets2 = new ActionReference();
            refTargets2.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("targetLayers"));
            refTargets2.putEnumerated(charIDToTypeID("Dcmn"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
            var descTargets2 = executeActionGet(refTargets2);

            if (descTargets2.hasKey(stringIDToTypeID("targetLayers"))) {
                var listTargets2 = descTargets2.getList(stringIDToTypeID("targetLayers"));
                for (var i = 0; i < listTargets2.count; i++) {
                    var refItem2 = listTargets2.getReference(i);
                    try {
                        addId(refItem2.getIdentifier());
                    } catch (e) {
                        var itemIndex2 = refItem2.getIndex();
                        var refByIndex2 = new ActionReference();
                        refByIndex2.putIndex(charIDToTypeID("Lyr "), itemIndex2);
                        var descByIndex2 = executeActionGet(refByIndex2);
                        if (descByIndex2.hasKey(stringIDToTypeID("layerID"))) {
                            addId(descByIndex2.getInteger(stringIDToTypeID("layerID")));
                        }
                    }
                }
            }
        } catch (e) {}

        try {
            if (doc.activeLayer) addId(doc.activeLayer.id);
        } catch (e) {}

        return ids;
    }

    function getSelectedLayers() {
        var selectedLayers = [];
        var ids = getSelectedLayerIDs();

        for (var i = 0; i < ids.length; i++) {
            var lyr = findLayerByNameOrId(doc, ids[i]);
            if (lyr) selectedLayers.push(lyr);
        }

        if (selectedLayers.length === 0) {
            try {
                if (doc.activeLayer) selectedLayers.push(doc.activeLayer);
            } catch (e) {}
        }

        return selectedLayers;
    }

    // 辅助函数：通过ID找到DOM图层对象
    function findLayerByNameOrId(container, id) {
        for (var i = 0; i < container.layers.length; i++) {
            var layer = container.layers[i];
            try {
                if (layer.id === id) return layer;
            } catch(e) {}
            if (layer.typename === "LayerSet") {
                var found = findLayerByNameOrId(layer, id);
                if (found) return found;
            }
        }
        return null;
    }

    // 根据规则计算新名称
    function calculateNewName(oldName, mode, param1, param2) {
        switch (mode) {
            case 0: // 在前面添加
                return param1 + oldName;
            case 1: // 在后面添加
                return oldName + param1;
            case 2: // 查找替换
                // 全局替换纯文本
                return oldName.split(param1).join(param2);
            case 3: // 正则查找替换
                try {
                    var reg = new RegExp(param1, "g");
                    return oldName.replace(reg, param2);
                } catch (e) {
                    return "[正则错误] " + oldName;
                }
            default:
                return oldName;
        }
    }

    function getLayerNameByID(layerID) {
        var ref = new ActionReference();
        ref.putIdentifier(charIDToTypeID("Lyr "), layerID);
        var desc = executeActionGet(ref);
        return desc.getString(charIDToTypeID("Nm  "));
    }

    function setLayerNameByID(layerID, newName) {
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putIdentifier(charIDToTypeID("Lyr "), layerID);
        desc.putReference(charIDToTypeID("null"), ref);

        var descLayer = new ActionDescriptor();
        descLayer.putString(charIDToTypeID("Nm  "), newName);
        desc.putObject(charIDToTypeID("T   "), charIDToTypeID("Lyr "), descLayer);

        executeAction(charIDToTypeID("setd"), desc, DialogModes.NO);
    }

    function showCompletionDialog(targetCount, changedCount, alertReport) {
        var dlg = new Window("dialog", "完成通知");
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.spacing = 10;
        dlg.margins = 16;

        dlg.add("statictext", undefined, "图层重命名完成！");
        dlg.add("statictext", undefined, "共处理了 " + targetCount + " 个图层，实际修改了 " + changedCount + " 个图层。");

        var maxLines = 500;
        var originalCount = alertReport.length;
        if (alertReport.length > maxLines) {
            alertReport = alertReport.slice(0, maxLines);
            alertReport.push("...（已截断，仅显示前 " + maxLines + " 条，共 " + originalCount + " 条）");
        }

        dlg.add("statictext", undefined, "详细变更历史：");
        var txt = dlg.add("edittext", undefined, alertReport.join("\r\n"), { multiline: true, scrolling: true });
        txt.preferredSize = [600, 360];
        try { txt.readonly = true; } catch (e) {}

        var btnGroup = dlg.add("group");
        btnGroup.alignment = "right";
        btnGroup.add("button", undefined, "关闭", { name: "ok" });

        dlg.center();
        dlg.show();
    }

    function getSelectedTargets() {
        var targets = [];
        var ids = getSelectedLayerIDs();

        if (ids.length > 0) {
            for (var i = 0; i < ids.length; i++) {
                targets.push({ kind: "id", id: ids[i] });
            }
            return targets;
        }

        try {
            if (doc.activeLayer) targets.push({ kind: "dom", layer: doc.activeLayer });
        } catch (e) {}

        return targets;
    }

    // 获取目标图层（带用户确认逻辑）
    function getTargetLayers() {
        var targets = getSelectedTargets();
        // PS 有时候会默认选中背景层，判断真正多选或单选
        if (targets.length === 0 || (targets.length === 1 && targets[0].kind === "dom" && targets[0].layer === doc.backgroundLayer && doc.layers.length > 1)) {
            var confirmAll = confirm("当前未检测到选中的图层，是否对【整个文档的所有图层】生效？", false, "确认操作范围");
            if (confirmAll) {
                var allLayers = getAllLayers(doc);
                var allTargets = [];
                for (var i = 0; i < allLayers.length; i++) {
                    allTargets.push({ kind: "dom", layer: allLayers[i] });
                }
                return allTargets;
            } else {
                return null;
            }
        }
        return targets;
    }

    // --- 按钮事件绑定 ---

    // 检查（预演）按钮
    btnCheck.onClick = function() {
        var targets = getTargetLayers();
        if (!targets || targets.length === 0) return;

        var mode = dropDown.selection.index;
        var p1 = txtIn1.text;
        var p2 = txtIn2.text;
        
        var report = [];
        for (var i = 0; i < targets.length; i++) {
            var oldName = targets[i].kind === "dom" ? targets[i].layer.name : getLayerNameByID(targets[i].id);
            var newName = calculateNewName(oldName, mode, p1, p2);
            report.push("原名: " + oldName + "  ->  新名: " + newName);
        }
        
        txtResult.text = report.join("\r\n");
    };

    // 重命名（确认执行）按钮
    btnRename.onClick = function() {
        var targets = getTargetLayers();
        if (!targets || targets.length === 0) return;

        var mode = dropDown.selection.index;
        var p1 = txtIn1.text;
        var p2 = txtIn2.text;

        var changedCount = 0;
        var alertReport = [];

        // 开启历史记录记录，方便一次性 Undo
        app.activeDocument.suspendHistory("批量重命名图层", "doRename()");

        function doRename() {
            for (var i = 0; i < targets.length; i++) {
                var oldName = targets[i].kind === "dom" ? targets[i].layer.name : getLayerNameByID(targets[i].id);
                var newName = calculateNewName(oldName, mode, p1, p2);

                if (oldName !== newName) {
                    if (targets[i].kind === "dom") {
                        targets[i].layer.name = newName;
                    } else {
                        setLayerNameByID(targets[i].id, newName);
                    }
                    changedCount++;
                }
                alertReport.push("Before: " + oldName + "  >>>  After: " + newName);
            }
        }

        win.close(); // 关闭当前对话框

        showCompletionDialog(targets.length, changedCount, alertReport);
    };

    // 取消按钮
    btnCancel.onClick = function() {
        win.close();
    };

    win.center();
    win.show();
}

main();
