// ==========================================================================
// PS_STATUS - Photoshop 状态层管理脚本
//
// 使用方法：
//   1. 选中 status 组内的状态层或状态组 -> 运行脚本 -> 批量更新选中的状态
//   2. 选中包含换行/Tab 的多行文本层（TSV 格式）-> 运行脚本 -> 批量创建状态层
//   3. 不选中任何图层 -> 运行脚本 -> 弹出菜单选择：
//      a) 一键更新全部状态层
//      b) 从当前可见状态反推配置，新建/更新状态层
//      c) 从 Excel 粘贴 TSV，一键创建+更新一批状态层
//
// 概念说明：
//   - key 组：名称以英文冒号结尾的图层组（如 "color:"），代表一个可切换属性
//   - key 文本层：名称以英文冒号结尾的文本层，值会被直接写入文本内容
//   - status 组：名称为 "status" 的图层组，存放状态层
//   - 状态层：名称格式为 "Key1:Value1,Key2:Value2" 的普通图层
// =============================================================================

#target photoshop

var TRUE_VALUES = ["true", "1", "yes", "on", "显示"];
var FALSE_VALUES = ["false", "0", "no", "off", "隐藏"];
var EMPTY_VALUES = ["", "-", "/", "\\", "empty", "null"];

var stepDebug = false;
var debug = false;
var lastParseTSVDebugText = "";
var psStatusExcelConfigByLayerId = {};
var psStatusExcelDebugHeaderSetByLayerId = {};
var PS_STATUS_MANUAL_HIDDEN_SUFFIX = " __PS_STATUS_MANUAL_HIDDEN";
var PS_STATUS_SHOW_PARENTS_POLICY = "show_parent_chain";


if (app.documents.length > 0) app.activeDocument.suspendHistory("PS_STATUS", "psStatusRunSafe()");


function debugStep(msg) {
    if (stepDebug) {
        if (!confirm("Debug 步骤: " + msg + "\n\n确定=继续，取消=退出")) {
            throw new Error("DEBUG_EXIT");
        }
    }
}

function debugConfirmLines(title, lines) {
    if (!stepDebug) return;
    var pageSize = 18;
    var idx = 0;
    while (idx < lines.length) {
        var end = Math.min(idx + pageSize, lines.length);
        var slice = [];
        for (var i = idx; i < end; i++) slice.push(lines[i]);
        var msg = title + "\n(" + (idx + 1) + "-" + end + "/" + lines.length + ")\n\n" + slice.join("\n");
        if (!confirm(msg + "\n\n确定=下一页/继续，取消=退出")) throw new Error("DEBUG_EXIT");
        idx = end;
    }
}

function debugDialogText(title, text) {
    if (!debug) return;
    var dlg = new Window("dialog", String(title || "Debug"));
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var box = dlg.add("edittext", undefined, String(text || ""), { multiline: true, scrolling: true });
    box.preferredSize = [760, 360];
    try { box.active = true; } catch (e) {}

    var btns = dlg.add("group");
    btns.alignment = "right";
    btns.add("button", undefined, "关闭", { name: "ok" });

    dlg.show();
}

function dialogContinueOrExit(title, text) {
    var dlg = new Window("dialog", String(title || "提示"));
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var box = dlg.add("edittext", undefined, String(text || ""), { multiline: true, scrolling: true });
    box.preferredSize = [760, 360];
    try { box.active = true; } catch (e) {}

    var btns = dlg.add("group");
    btns.alignment = "right";
    var btnExit = btns.add("button", undefined, "退出");
    var btnContinue = btns.add("button", undefined, "继续");

    var ok = false;
    btnContinue.onClick = function () { ok = true; dlg.close(1); };
    btnExit.onClick = function () { ok = false; dlg.close(0); };

    dlg.show();
    return ok;
}




function psStatusRunSafe() {
    try {
        psStatusRun();
    } catch(e) {
        if (e.message !== "DEBUG_EXIT") {
            alert("脚本执行出错: " + e.message + "\nLine: " + e.line);
        }
    }
}

function psStatusRun() {
    debugStep("psStatusRun 开始");
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    var selectedLayers = getSelectedLayers(doc);
    debugStep("获取选中图层: " + selectedLayers.length);
    var statusGroups = psStatusFindStatusGroups(doc);
    debugStep("获取 status groups: " + statusGroups.length);

    if (selectedLayers.length === 0) {
        var pick = psStatusShowNoSelectionDialog();
        debugStep("选择模式: " + pick);
        if (pick === "mode1") {
            psStatusRunAutoUpdate(doc, statusGroups);
            return;
        }
        if (pick === "mode4") {
            psStatusRunCreateFromCurrent(doc);
            return;
        }
        if (pick === "mode5") {
            psStatusRunExcelPalette(doc);
            return;
        }
        return;
    }

    var tasks = psStatusCollectTasksFromSelection(selectedLayers, statusGroups);
    debugStep("获取 tasks: " + tasks.length);
    if (tasks.length > 0) {
        psStatusRunUpdateTasks(doc, tasks);
        return;
    }

    var textLayers = psStatusCollectSelectedTextLayers(selectedLayers);
    debugStep("获取 text layers: " + textLayers.length);
    if (textLayers.length > 0) {
        psStatusRunMultiLineText(doc, textLayers);
        return;
    }

    alert(
        "当前选中的图层未匹配到可执行任务。\n\n" +
        "可用的两种选中模式：\n" +
        "1) 选中 status 组内的状态层或状态组（批量更新选中）\n" +
        "2) 选中包含换行/Tab 的文本层（批量创建状态层）\n\n" +
        "如果要一键更新全部状态层：请取消选中图层后运行脚本，并在弹窗中选择“更新全部状态层”。"
    );
    return;
}

function psStatusRunAutoUpdate(doc, statusGroups) {
    var tasks = psStatusCollectAllStatusLayers(statusGroups);
    if (tasks.length === 0) {
        var pick = psStatusShowNoSelectionDialog();
        if (pick === "mode1") {
            alert("未找到任何状态层可更新。");
            return;
        }
        if (pick === "mode4") {
            psStatusRunCreateFromCurrent(doc);
            return;
        }
        if (pick === "mode5") {
            psStatusRunExcelPalette(doc);
            return;
        }
        return;
    }
    psStatusRunUpdateTasks(doc, tasks);
}

function psStatusRunUpdateTasks(doc, tasks) {
    var allKeyTargets = psStatusCollectKeyTargetsExcludingStatus(doc);
    var targetsByKey = psStatusIndexTargetsByKey(allKeyTargets);
    var allStatusGroups = psStatusFindStatusGroups(doc);

    function psStatusGetLayerPath(layer) {
        if (!layer) return "";
        var names = [];
        var cur = layer;
        while (cur) {
            names.push(String(cur.name || ""));
            if (!cur.parent || cur.parent.typename === "Document") break;
            cur = cur.parent;
        }
        names.reverse();
        return names.join(" / ");
    }

    function psStatusGetFirstHiddenParent(layer) {
        var p = layer ? layer.parent : null;
        while (p && p.typename !== "Document") {
            if (!getVisibleSafe(p)) return p;
            p = p.parent;
        }
        return null;
    }

    function psStatusCollectWouldShowLayersForConfig(configMap, pureKey, valueText, targets) {
        var res = [];
        var lower = valueText.toLowerCase();
        var isEmpty = containsValue(EMPTY_VALUES, lower) || valueText === "";
        var isTrue = containsValue(TRUE_VALUES, lower);
        var isFalse = containsValue(FALSE_VALUES, lower);

        for (var i = 0; i < targets.length; i++) {
            var target = targets[i];
            if (psStatusIsKeyTextLayer(target)) {
                if (!isEmpty) res.push(target);
                continue;
            }
            if (!target || target.typename !== "LayerSet") continue;

            var gt = psStatusDetectKeyGroupType(target, pureKey);
            if (gt.type === "textContent") {
                if (isEmpty) continue;
                for (var b = 0; b < gt.keyNameTextLayers.length; b++) res.push(gt.keyNameTextLayers[b]);
                for (var c = 0; c < target.layers.length; c++) {
                    var child = target.layers[c];
                    if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
                    res.push(child);
                }
                continue;
            }

            if (gt.type === "toggle") {
                if (isFalse || isEmpty) continue;
                for (var d = 0; d < gt.keyNameChildren.length; d++) res.push(gt.keyNameChildren[d]);
                continue;
            }

            if (isEmpty) continue;
            var desired = splitCommaValues(valueText);
            for (var g = 0; g < target.layers.length; g++) {
                var v = target.layers[g];
                if (psStatusIsKeyGroup(v) || psStatusIsKeyTextLayer(v)) continue;
                if (arrayContains(desired, trim(String(v.name)))) res.push(v);
            }
        }
        return res;
    }

    function psStatusBuildHideReasonMapForConfig(configMap, targetsByKey) {
        var reasonMap = {};

        function addReason(layer, key, value, mode) {
            var id = getLayerIdSafe(layer);
            if (id === null) return;
            var sid = String(id);
            if (!reasonMap.hasOwnProperty(sid)) reasonMap[sid] = [];
            reasonMap[sid].push({ key: String(key), value: String(value), mode: String(mode || "") });
        }

        for (var key in configMap) {
            if (!configMap.hasOwnProperty(key)) continue;
            var pureKey = trim(String(key));
            if (!pureKey) continue;
            if (pureKey.indexOf("__") === 0) continue;
            var valueText = trim(String(configMap[key]));
            var lower = valueText.toLowerCase();
            var isEmpty = containsValue(EMPTY_VALUES, lower) || valueText === "";
            var isTrue = containsValue(TRUE_VALUES, lower);
            var isFalse = containsValue(FALSE_VALUES, lower);
            var targets = targetsByKey.hasOwnProperty(pureKey) ? targetsByKey[pureKey] : [];

            for (var i = 0; i < targets.length; i++) {
                var target = targets[i];
                if (psStatusIsKeyTextLayer(target)) {
                    if (isEmpty) addReason(target, pureKey, valueText, "keyTextLayer(empty)");
                    continue;
                }
                if (!target || target.typename !== "LayerSet") continue;

                var gt = psStatusDetectKeyGroupType(target, pureKey);
                if (gt.type === "textContent") {
                    if (!isEmpty) continue;
                    for (var a = 0; a < target.layers.length; a++) addReason(target.layers[a], pureKey, valueText, "textContent(empty)");
                    continue;
                }

                if (gt.type === "toggle") {
                    if (!(isFalse || isEmpty)) continue;
                    for (var b = 0; b < gt.keyNameChildren.length; b++) addReason(gt.keyNameChildren[b], pureKey, valueText, "toggle(off)");
                    continue;
                }

                if (isEmpty) {
                    for (var c = 0; c < target.layers.length; c++) {
                        var child0 = target.layers[c];
                        if (psStatusIsKeyGroup(child0) || psStatusIsKeyTextLayer(child0)) continue;
                        addReason(child0, pureKey, valueText, "select(allHidden)");
                    }
                    continue;
                }

                var desired = splitCommaValues(valueText);
                for (var d = 0; d < target.layers.length; d++) {
                    var child = target.layers[d];
                    if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
                    var show = arrayContains(desired, trim(String(child.name)));
                    if (!show) addReason(child, pureKey, valueText, "select(notSelected)");
                }
            }
        }
        return reasonMap;
    }

    function psStatusCollectShowParentConflicts(validated, targetsByKey) {
        var hiddenParents = [];
        var overriddenParents = [];
        var seenA = {};
        var seenB = {};

        for (var i = 0; i < validated.length; i++) {
            var statusLayer = validated[i].layer;
            var configMap = validated[i].configMap || {};
            var hideReasonMap = psStatusBuildHideReasonMapForConfig(configMap, targetsByKey);

            for (var key in configMap) {
                if (!configMap.hasOwnProperty(key)) continue;
                var pureKey = trim(String(key));
                if (!pureKey) continue;
                if (pureKey.indexOf("__") === 0) continue;
                var valueText = trim(String(configMap[key]));
                var targets = targetsByKey.hasOwnProperty(pureKey) ? targetsByKey[pureKey] : [];
                var wouldShow = psStatusCollectWouldShowLayersForConfig(configMap, pureKey, valueText, targets);

                for (var w = 0; w < wouldShow.length; w++) {
                    var layerToShow = wouldShow[w];

                    var hiddenParent = psStatusGetFirstHiddenParent(layerToShow);
                    if (hiddenParent) {
                        var id1 = getLayerIdSafe(layerToShow);
                        var id2 = getLayerIdSafe(hiddenParent);
                        if (id1 !== null && id2 !== null) {
                            var k2 = String(id1) + "|" + String(id2);
                            if (!seenA.hasOwnProperty(k2)) {
                                seenA[k2] = true;
                                hiddenParents.push({
                                    statusLayer: statusLayer,
                                    key: pureKey,
                                    value: valueText,
                                    showLayer: layerToShow,
                                    hiddenParent: hiddenParent
                                });
                            }
                        }
                    }

                    var p = layerToShow ? layerToShow.parent : null;
                    while (p && p.typename !== "Document") {
                        var pid = getLayerIdSafe(p);
                        if (pid !== null && hideReasonMap.hasOwnProperty(String(pid))) {
                            var sid = getLayerIdSafe(layerToShow);
                            if (sid !== null) {
                                var k3 = String(sid) + "|" + String(pid);
                                if (!seenB.hasOwnProperty(k3)) {
                                    seenB[k3] = true;
                                    overriddenParents.push({
                                        statusLayer: statusLayer,
                                        key: pureKey,
                                        value: valueText,
                                        showLayer: layerToShow,
                                        hiddenAncestor: p,
                                        reasons: hideReasonMap[String(pid)]
                                    });
                                }
                            }
                        }
                        p = p.parent;
                    }
                }
            }
        }
        return { hiddenParents: hiddenParents, overriddenParents: overriddenParents };
    }

    function psStatusShowParentConflictDialog(allIssueBlocks, issueCount, parentConflictsA, parentConflictsB) {
        var lines = [];
        if (allIssueBlocks && allIssueBlocks.length > 0) {
            lines.push("【配置校验问题】共 " + (issueCount || 0) + " 处（忽略也可继续执行）：");
            lines.push("");
            for (var i = 0; i < allIssueBlocks.length; i++) {
                lines.push(allIssueBlocks[i]);
                lines.push("");
            }
        }
        if (parentConflictsA && parentConflictsA.length > 0) {
            lines.push("【父组可见性冲突】共 " + parentConflictsA.length + " 处：");
            lines.push("下列图层需要点亮父组才能显示，但其父组当前处于隐藏状态。");
            lines.push("你可以选择：点亮直接父组 / 点亮父组链 / 保持父组现状 / 选中问题图层并退出。");
            lines.push("");
            for (var j = 0; j < parentConflictsA.length; j++) {
                var c = parentConflictsA[j];
                var sName = String(c.statusLayer && c.statusLayer.name ? c.statusLayer.name : "(unnamed)");
                var showPath = psStatusGetLayerPath(c.showLayer);
                var parentPath = psStatusGetLayerPath(c.hiddenParent);
                lines.push("- 状态层: " + sName);
                lines.push("  key=" + c.key + "  value=" + c.value);
                lines.push("  需要显示: " + showPath);
                lines.push("  隐藏父组: " + parentPath);
                lines.push("");
            }
        }
        if (parentConflictsB && parentConflictsB.length > 0) {
            lines.push("【父组链覆盖冲突】共 " + parentConflictsB.length + " 处：");
            lines.push("下列图层即使点亮父组链，也可能被其他 key/value 隐藏父组分支而最终不可见。");
            lines.push("注意：key/value 的显隐结果优先级高于“点亮父组链”辅助显示；点亮父组链不会强行覆盖 key/value 的隐藏决定。");
            lines.push("");
            for (var k = 0; k < parentConflictsB.length; k++) {
                var c2 = parentConflictsB[k];
                var sName2 = String(c2.statusLayer && c2.statusLayer.name ? c2.statusLayer.name : "(unnamed)");
                var showPath2 = psStatusGetLayerPath(c2.showLayer);
                var ancPath = psStatusGetLayerPath(c2.hiddenAncestor);
                lines.push("- 状态层: " + sName2);
                lines.push("  key=" + c2.key + "  value=" + c2.value);
                lines.push("  需要显示: " + showPath2);
                lines.push("  可能被隐藏的父组/分支: " + ancPath);
                if (c2.reasons && c2.reasons.length > 0) {
                    for (var r = 0; r < c2.reasons.length; r++) {
                        lines.push("  隐藏原因: key=" + c2.reasons[r].key + "  value=" + c2.reasons[r].value + "  via=" + c2.reasons[r].mode);
                    }
                }
                lines.push("");
            }
        }

        var dlg = new Window("dialog", "预检查确认");
        dlg.orientation = "column";
        dlg.alignChildren = "fill";

        var box = dlg.add("edittext", undefined, lines.join("\n"), { multiline: true, scrolling: true,wantReturn: true, readonly: true });
        box.preferredSize = [860, 420];
        try { box.active = true; } catch (e) {}

        var btns = dlg.add("group");
        btns.alignment = "right";
        var bParentOnly = btns.add("button", undefined, "点亮直接父组并继续");
        var bChain = btns.add("button", undefined, "点亮父组链并继续");
        var bKeep = btns.add("button", undefined, "保持父组现状并继续");
        var bPick = btns.add("button", undefined, "选中问题图层并退出");

        var choice = null;
        bParentOnly.onClick = function () { choice = "show_parent_only"; dlg.close(1); };
        bChain.onClick = function () { choice = "show_parent_chain"; dlg.close(1); };
        bKeep.onClick = function () { choice = "keep_as_now"; dlg.close(1); };
        bPick.onClick = function () { choice = "select_and_exit"; dlg.close(1); };

        if (dlg.show() !== 1) return null;
        return choice;
    }

    function psStatusShowIssuesOnlyDialog(allIssueBlocks, issueCount) {
        var lines = [];
        lines.push("【配置校验问题】共 " + (issueCount || 0) + " 处：");
        lines.push("这些问题不会阻止执行，但可能导致部分 key/value 无效。");
        lines.push("");
        if (allIssueBlocks && allIssueBlocks.length > 0) {
            for (var i = 0; i < allIssueBlocks.length; i++) {
                lines.push(allIssueBlocks[i]);
                lines.push("");
            }
        }

        var dlg = new Window("dialog", "校验问题");
        dlg.orientation = "column";
        dlg.alignChildren = "fill";

        var box = dlg.add("edittext", undefined, lines.join("\n"), { multiline: true, scrolling: true, readonly: true });
        box.preferredSize = [860, 420];
        try { box.active = true; } catch (e) {}

        var btns = dlg.add("group");
        btns.alignment = "right";
        var bContinue = btns.add("button", undefined, "继续");
        var bPick = btns.add("button", undefined, "选中问题图层并退出");

        var choice = null;
        bContinue.onClick = function () { choice = "continue"; dlg.close(1); };
        bPick.onClick = function () { choice = "select_and_exit"; dlg.close(1); };

        if (dlg.show() !== 1) return null;
        return choice;
    }

    function psStatusCollectManualHiddenTextContentChildren(allKeyTargets) {
        var res = [];
        var seen = {};
        for (var i = 0; i < allKeyTargets.length; i++) {
            var t = allKeyTargets[i];
            if (!t || t.typename !== "LayerSet") continue;
            var pureKey = psStatusGetPureKeyName(t.name);
            var gt = psStatusDetectKeyGroupType(t, pureKey);
            if (!gt || gt.type !== "textContent") continue;

            for (var j = 0; j < t.layers.length; j++) {
                var child = t.layers[j];
                if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
                var id = getLayerIdSafe(child);
                if (id === null) continue;
                var sid = String(id);
                if (seen.hasOwnProperty(sid)) continue;
                seen[sid] = true;
                if (!getVisibleSafe(child)) res.push({ layer: child, id: sid, originalName: String(child.name || "") });
            }
        }
        return res;
    }

    function psStatusMarkManualHidden(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            try {
                if (!it || !it.layer) continue;
                if (String(it.layer.name || "").indexOf(PS_STATUS_MANUAL_HIDDEN_SUFFIX) >= 0) continue;
                it.layer.name = String(it.originalName || "") + PS_STATUS_MANUAL_HIDDEN_SUFFIX;
            } catch (e) {}
        }
    }

    function psStatusRestoreManualHiddenBeforeFlatten(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            try {
                if (!it || !it.layer) continue;
                if (String(it.layer.name || "").indexOf(PS_STATUS_MANUAL_HIDDEN_SUFFIX) < 0) continue;
                it.layer.name = String(it.originalName || "");
                trySetVisible(it.layer, false);
            } catch (e) {}
        }
    }

    function psStatusUnmarkManualHidden(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            try { if (it && it.layer) it.layer.name = String(it.originalName || ""); } catch (e) {}
        }
    }

    var manualHiddenItems = psStatusCollectManualHiddenTextContentChildren(allKeyTargets);
    psStatusMarkManualHidden(manualHiddenItems);

    try {
    var prevShowParents = PS_STATUS_SHOW_PARENTS_POLICY;
    PS_STATUS_SHOW_PARENTS_POLICY = "show_parent_chain";

    function psStatusExcelGetStoredConfig(layer) {
        var id = getLayerIdSafe(layer);
        if (id === null) return null;
        var k = String(id);
        if (!psStatusExcelConfigByLayerId.hasOwnProperty(k)) return null;
        return {
            configMap: psStatusExcelConfigByLayerId[k],
            debugHeaderSet: psStatusExcelDebugHeaderSetByLayerId.hasOwnProperty(k) ? psStatusExcelDebugHeaderSetByLayerId[k] : null
        };
    }

    function psStatusHideDebugTargetsNotInHeader(allKeyTargets, headerSet) {
        for (var i = 0; i < allKeyTargets.length; i++) {
            var t = allKeyTargets[i];
            var pure = psStatusGetPureKeyName(t && t.name ? t.name : "");
            if (pure.indexOf("__") !== 0) continue;
            if (headerSet && headerSet.hasOwnProperty(pure)) continue;
            trySetVisible(t, false);
        }
    }

    function psStatusHideAllDebugTargets(allKeyTargets) {
        psStatusHideDebugTargetsNotInHeader(allKeyTargets, {});
    }

    var validated = [];
    var allIssueBlocks = [];
    var allIssueCount = 0;
    var issueTargetIds = [];
    var issueTargetIdSet = {};
    for (var p = 0; p < tasks.length; p++) {
        var statusLayerP = tasks[p];
        var storedP = psStatusExcelGetStoredConfig(statusLayerP);
        var configMapP = storedP ? storedP.configMap : psStatusParseStatusLayerConfig(statusLayerP.name);
        if (!configMapP) continue;
        validated.push({ layer: statusLayerP, configMap: configMapP, debugHeaderSet: storedP ? storedP.debugHeaderSet : null });
        var resultP = psStatusValidateConfig(configMapP, targetsByKey); var issuesP = resultP.issues;
        if (issuesP.length > 0) {
            var block = String(statusLayerP.name || "(unnamed)") + "\n" + issuesP.join("\n");
            allIssueBlocks.push(block);
            allIssueCount += resultP.issueCount || issuesP.length;
            var targetIdsP = resultP.targetIds;
            for (var ti = 0; ti < targetIdsP.length; ti++) {
                var tid = targetIdsP[ti];
                if (tid === null) continue;
                var stid = String(tid);
                if (issueTargetIdSet.hasOwnProperty(stid)) continue;
                issueTargetIdSet[stid] = true;
                issueTargetIds.push(tid);
            }
        }
    }
    var parentConflicts = psStatusCollectShowParentConflicts(validated, targetsByKey);
    var hasParentConflicts = parentConflicts && ((parentConflicts.hiddenParents && parentConflicts.hiddenParents.length > 0) || (parentConflicts.overriddenParents && parentConflicts.overriddenParents.length > 0));
    if (hasParentConflicts) {
        var pick = psStatusShowParentConflictDialog(allIssueBlocks, allIssueCount, parentConflicts.hiddenParents || [], parentConflicts.overriddenParents || []);
        if (!pick) return;
        if (pick === "select_and_exit") {
            var ids = [];
            var seen2 = {};
            var listA = parentConflicts.hiddenParents || [];
            for (var cc = 0; cc < listA.length; cc++) {
                var pid = getLayerIdSafe(listA[cc].hiddenParent);
                if (pid === null) continue;
                var spid = String(pid);
                if (seen2.hasOwnProperty(spid)) continue;
                seen2[spid] = true;
                ids.push(pid);
            }
            var listB = parentConflicts.overriddenParents || [];
            for (var dd = 0; dd < listB.length; dd++) {
                var pid2 = getLayerIdSafe(listB[dd].hiddenAncestor);
                if (pid2 === null) continue;
                var spid2 = String(pid2);
                if (seen2.hasOwnProperty(spid2)) continue;
                seen2[spid2] = true;
                ids.push(pid2);
            }
            for (var ee = 0; ee < issueTargetIds.length; ee++) {
                var gid2 = issueTargetIds[ee];
                if (gid2 === null) continue;
                var sgid2 = String(gid2);
                if (seen2.hasOwnProperty(sgid2)) continue;
                seen2[sgid2] = true;
                ids.push(gid2);
            }
            try { selectLayersByIds(doc, ids); } catch (e2) { if (ids.length > 0) try { selectLayer(doc, findLayerByID(doc, ids[0])); } catch (e3) {} }
            return;
        }
        if (pick === "show_anyway") pick = "show_parent_chain";
        PS_STATUS_SHOW_PARENTS_POLICY = pick;
    } else if (allIssueBlocks.length > 0) {
        var pick2 = psStatusShowIssuesOnlyDialog(allIssueBlocks, allIssueCount);
        if (!pick2) return;
        if (pick2 === "select_and_exit") {
            var groupIds2 = [];
            var seenGroup2 = {};
            for (var ff = 0; ff < issueTargetIds.length; ff++) {
                var gidQ = issueTargetIds[ff];
                if (gidQ === null) continue;
                var sgidQ = String(gidQ);
                if (seenGroup2.hasOwnProperty(sgidQ)) continue;
                seenGroup2[sgidQ] = true;
                groupIds2.push(gidQ);
            }
            try { selectLayersByIds(doc, groupIds2); } catch (e4) { if (groupIds2.length > 0) try { selectLayer(doc, findLayerByID(doc, groupIds2[0])); } catch (e5) {} }
            return;
        }
    }

    var updated = 0;
    for (var i = 0; i < validated.length; i++) {
        var statusLayer = validated[i].layer;
        var configMap = validated[i].configMap;
        var debugHeaderSet = validated[i].debugHeaderSet;
        psStatusResetAllKeyTargets(allKeyTargets);
        psStatusHideDebugTargetsNotInHeader(allKeyTargets, debugHeaderSet);
        psStatusApplyConfigByDepth(targetsByKey, configMap);

        psStatusRestoreManualHiddenBeforeFlatten(manualHiddenItems);
        var newLayer = psStatusFillStatusWithFlatten(doc, statusLayer, allStatusGroups);
        psStatusMarkManualHidden(manualHiddenItems);
        psStatusHideAllDebugTargets(allKeyTargets);
        if (newLayer) updated++;
    }

    app.refresh();
    alert("处理完成！共更新 " + updated + " 个状态。");
    } finally {
        PS_STATUS_SHOW_PARENTS_POLICY = prevShowParents;
        psStatusUnmarkManualHidden(manualHiddenItems);
    }
}

function psStatusRunMultiLineText(doc, textLayers) {
    var createdLayers = [];
    for (var i = 0; i < textLayers.length; i++) {
        var tl = textLayers[i];
        var content = "";
        try { content = String(tl.textItem.contents || ""); } catch (e) { content = ""; }
        if (!content) continue;

        var container = tl.parent;
        var statusGroup = psStatusGetOrCreateStatusGroup(container);

        if (content.indexOf("\t") >= 0) {
            var groupName = "excel_from_text";
            var created = psStatusCreateFromExcelText(doc, statusGroup, groupName, content);
            createdLayers = createdLayers.concat(created);
        } else {
            var lines = splitLines(content);
            for (var k = 0; k < lines.length; k++) {
                var line = trim(lines[k]);
                if (!line) continue;
                var layerName = psStatusNormalizeMultiLineStatusName(line, k + 1);
                var layer = statusGroup.artLayers.add();
                layer.name = layerName;
                createdLayers.push(layer);
            }
        }
    }

    if (createdLayers.length === 0) {
        alert("未从文本层创建到任何状态层。");
        return;
    }

    psStatusRunUpdateTasks(doc, createdLayers);
}

function psStatusRunCreateFromCurrent(doc) {
    var statusGroup = psStatusGetTopStatusGroupOrCreate(doc);
    if (!statusGroup) {
        alert("错误：未找到顶层唯一的 status group。");
        return;
    }

    var allKeyTargets = psStatusCollectKeyTargetsExcludingStatus(doc);
    var stateResult = psStatusBuildConfigFromCurrentState(allKeyTargets);
    if (stateResult.issues.length > 0) {
        var msg = "发现以下问题：\n\n" + stateResult.issues.join("\n") + "\n\n是否忽略并继续？";
        if (!dialogContinueOrExit("校验问题", msg)) return;
    }

    var normalized = psStatusNormalizeConfigMap(stateResult.configMap);
    var existing = psStatusFindExistingStatusLayerByConfig(statusGroup, normalized);
    var statusLayer;

    if (existing) {
        statusLayer = existing;
    } else {
        statusLayer = statusGroup.artLayers.add();
        statusLayer.name = "status={}";
    }

    statusLayer.name = "status={" + psStatusSerializeConfigMap(stateResult.configMap, null) + "}";
    try { statusLayer.visible = true; } catch (e) {}
    try { statusLayer.move(statusGroup, ElementPlacement.PLACEATBEGINNING); } catch (e) {}

    var allStatusGroups = psStatusFindStatusGroups(doc);
    var targetsByKey = psStatusIndexTargetsByKey(allKeyTargets);

    var validateResult2 = psStatusValidateConfig(stateResult.configMap, targetsByKey);
    var issues2 = validateResult2.issues;
    if (issues2.length > 0) {
        var msg2 = "发现以下问题：\n\n" + issues2.join("\n") + "\n\n是否忽略并继续？";
        if (!dialogContinueOrExit("校验问题", msg2)) return;
    }

    psStatusResetAllKeyTargets(allKeyTargets);
    psStatusApplyConfigByDepth(targetsByKey, stateResult.configMap);
    var newLayer = psStatusFillStatusWithFlatten(doc, statusLayer, allStatusGroups);
    if (newLayer) {
        try { showLayerAndParents(newLayer); } catch (e) {}
        try { newLayer.move(newLayer.parent, ElementPlacement.PLACEATBEGINNING); } catch (e) {}
        selectLayer(doc, newLayer);
    }
}

function psStatusRunExcelPalette(doc) {
    debugStep("psStatusRunExcelPalette start");
    var statusGroup = psStatusGetTopStatusGroupOrCreate(doc);
    if (!statusGroup) {
        alert("错误：未找到顶层唯一的 status group。");
        return;
    }

    var pal = new Window("dialog", "PS_STATUS Excel");
    pal.orientation = "column";
    pal.alignChildren = "fill";

    var groupRow = pal.add("group");
    groupRow.orientation = "row";
    groupRow.alignChildren = "center";
    groupRow.add("statictext", undefined, "group 名称：");
    var groupNameInput = groupRow.add("edittext", undefined, "excel_status");
    groupNameInput.characters = 24;

    var input = pal.add("edittext", undefined, "", { multiline: true, scrolling: true });
    input.preferredSize = [720, 360];

    var btnRow = pal.add("group");
    btnRow.alignment = "right";
    var btnRun = btnRow.add("button", undefined, "创建并更新");
    var btnCancel = btnRow.add("button", undefined, "取消");

    var payload = { cancelled: true, groupName: "", text: "" };

    btnRun.onClick = function () {
        payload.cancelled = false;
        payload.groupName = trim(groupNameInput.text);
        payload.text = String(input.text || "");
        pal.close();
    };
    btnCancel.onClick = function () {
        payload.cancelled = true;
        pal.close();
    };

    debugStep("showing dialog");
    pal.show();
    debugStep("dialog closed");

    if (payload.cancelled) return;
    if (!payload.text) {
        alert("Excel 内容为空。");
        return;
    }

    try {
        var created = psStatusCreateFromExcelText(doc, statusGroup, payload.groupName || "excel_status", payload.text);
        debugStep("created excel layers");
        if (created.length === 0) {
            alert("未从 Excel 内容创建到任何状态层。");
            return;
        }

        psStatusRunUpdateTasks(doc, created);
        debugStep("psStatusRunExcelPalette done");
    } catch(e) {
        alert("psStatusRunExcelPalette 发生错误: " + e.message + "\nLine: " + e.line);
        throw e;
    }
}

function psStatusCreateFromExcelText(doc, statusGroup, groupName, excelText) {
    var rows = parseTSV(excelText);
    if (rows.length < 2) return [];

    var header = rows[0];
    var keys = [];
    var serializeKeys = [];
    var debugHeaderSet = {};
    for (var h = 1; h < header.length; h++) {
        var k = trim(header[h]);
        if (k) {
            keys.push(k);
            if (k.indexOf("__") !== 0) serializeKeys.push(k);
            else debugHeaderSet[k] = true;
        } else {
            keys.push("");
        }
    }

    var previewNames = [];
    var previewItems = [];
    for (var pr = 1; pr < rows.length; pr++) {
        var prow = rows[pr];
        if (!prow || prow.length === 0) continue;
        var prowName = trim(prow[0] || "");
        if (!prowName) continue;
        var pconfigMap = {};
        for (var pc = 1; pc < prow.length && pc <= keys.length; pc++) {
            var pkeyName = keys[pc - 1];
            if (!pkeyName) continue;
            var pcell = prow[pc] !== undefined ? String(prow[pc]) : "";
            pconfigMap[pkeyName] = encodeTextValue(trim(pcell));
        }
        var layerNamePreview = prowName + "={" + psStatusSerializeConfigMap(pconfigMap, serializeKeys) + "}";
        previewNames.push(layerNamePreview);
        previewItems.push({
            name: prowName,
            layerName: layerNamePreview,
            configMap: pconfigMap,
            headerFields: header.slice(0),
            rowFields: prow.slice(0)
        });
    }
    if (debug) {
        var previewText = (lastParseTSVDebugText ? (lastParseTSVDebugText + "\n\n") : "") + previewNames.join("\n");
        if (!psStatusShowExcelDryRunDialog(doc, previewItems, previewText)) throw new Error("DEBUG_EXIT");
    }

    var targetGroup = statusGroup.layerSets.add();
    targetGroup.name = groupName;
    try { targetGroup.move(statusGroup, ElementPlacement.PLACEATBEGINNING); } catch (e) {}

    var createdLayers = [];

    for (var r = 1; r < rows.length; r++) {
        var row = rows[r];
        if (!row || row.length === 0) continue;
        var rowName = trim(row[0] || "");
        if (!rowName) continue;

        var configMap = {};
        for (var c = 1; c < row.length && c <= keys.length; c++) {
            var keyName = keys[c - 1];
            if (!keyName) continue;
            var cell = row[c] !== undefined ? String(row[c]) : "";
            configMap[keyName] = encodeTextValue(trim(cell));
        }

        var layer = targetGroup.artLayers.add();
        layer.name = rowName + "={" + psStatusSerializeConfigMap(configMap, serializeKeys) + "}";
        var lid = getLayerIdSafe(layer);
        if (lid !== null) {
            psStatusExcelConfigByLayerId[String(lid)] = configMap;
            psStatusExcelDebugHeaderSetByLayerId[String(lid)] = debugHeaderSet;
        }
        createdLayers.push(layer);
    }

    psStatusReverseLayersInGroup(targetGroup);
    return createdLayers;
}

function psStatusReverseLayersInGroup(layerSet) {
    if (!layerSet || layerSet.typename !== "LayerSet") return;
    if (!layerSet.layers || layerSet.layers.length < 2) return;

    var layers = [];
    for (var i = 0; i < layerSet.layers.length; i++) layers.push(layerSet.layers[i]);
    for (var j = layers.length - 1; j >= 0; j--) {
        try { layers[j].move(layerSet, ElementPlacement.PLACEATEND); } catch (e) {}
    }
}

function psStatusShowExcelDryRunDialog(doc, previewItems, previewText) {
    var SHOW_STRUCTURE = true;
    var STRUCTURE_SHOW_HIERARCHY = true;
    var STRUCTURE_PS_STATUS_ONLY = true;
    var STRUCTURE_VISIBLE_ONLY = false;
    var STRUCTURE_VISIBILITY_MARK = true;
    var STRUCTURE_LAYER_TYPE_MARK = true;
    var STRUCTURE_INDENT_USE_TAB = true;
    var STRUCTURE_INDENT_CUSTOM = "-";
    var STRUCTURE_SHOW_NOOP_ACTIONS = false;
    var STRUCTURE_TEXT_PREVIEW_MAXLEN = 80;
    var LIST_MAX_ITEMS = 250;
    var EXCEL_ESCAPE_NEWLINES = true;
    var UI_INFO_WIDTH = 1080;
    var UI_LEFT_WIDTH = 260;
    var UI_RIGHT_WIDTH = 820;
    var UI_LINE_HEIGHT = 18;
    var UI_TABLE_LINES = 4;
    var UI_TABLE_COL_GAP = 2;
    var UI_TABLE_USE_MONO_FONT = true;
    var UI_TABLE_MONO_FONT_NAME = "Courier New";
    var UI_LAYERNAME_EXTRA_LINES = 1;
    var UI_LAYERNAME_MAX_LINES = 12;
    var UI_LAYERNAME_BREAK_ON_COMMA = false;
    var UI_LAYERNAME_SPACE_BEFORE_KEY = true;

    if (!doc) return dialogContinueOrExit("Excel 解析预览", String(previewText || ""));

    var allKeyTargets = psStatusCollectKeyTargetsExcludingStatus(doc);
    var targetsByKey = psStatusIndexTargetsByKey(allKeyTargets);

    var dlg = new Window("dialog", "Excel 解析预览");
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var info = dlg.add("edittext", undefined, String(previewText || ""), { multiline: true, scrolling: true, readonly: true });
    info.preferredSize = [UI_INFO_WIDTH, 200];

    var body = dlg.add("group");
    body.orientation = "row";
    body.alignChildren = ["fill", "fill"];

    var left = body.add("group");
    left.orientation = "column";
    left.alignChildren = ["fill", "fill"];
    left.alignment = ["fill", "fill"];

    var list = left.add("listbox", undefined, [], { multiselect: false });
    list.preferredSize = [UI_LEFT_WIDTH, 420];
    list.alignment = ["fill", "fill"];

    var rightCol = body.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];
    rightCol.alignment = ["fill", "fill"];

    var rightExcelRaw = rightCol.add("edittext", undefined, "", { multiline: true, scrolling: true, readonly: true });
    rightExcelRaw.preferredSize = [UI_RIGHT_WIDTH, UI_TABLE_LINES * UI_LINE_HEIGHT + 0 ];
    rightExcelRaw.alignment = ["fill", "top"];
    rightExcelRaw.minimumSize = [UI_RIGHT_WIDTH, UI_TABLE_LINES * UI_LINE_HEIGHT];
    rightExcelRaw.maximumSize = [UI_RIGHT_WIDTH, UI_TABLE_LINES * UI_LINE_HEIGHT];

    var rightLayerName = rightCol.add("edittext", undefined, "", { multiline: true, scrolling: true, readonly: true });
    rightLayerName.preferredSize = [UI_RIGHT_WIDTH, (2 + UI_LAYERNAME_EXTRA_LINES) * UI_LINE_HEIGHT + 8];
    rightLayerName.alignment = ["fill", "top"];

    var rightStructure = rightCol.add("edittext", undefined, "", { multiline: true, scrolling: true, readonly: true });
    rightStructure.preferredSize = [UI_RIGHT_WIDTH, 220];
    rightStructure.alignment = ["fill", "fill"];
    rightStructure.minimumSize = [UI_RIGHT_WIDTH, 120];

    var btns = dlg.add("group");
    btns.alignment = "right";
    var btnContinue = btns.add("button", undefined, "继续");
    var btnExit = btns.add("button", undefined, "退出");

    var shownItems = [];
    for (var i = 0; i < (previewItems ? previewItems.length : 0); i++) {
        if (shownItems.length >= LIST_MAX_ITEMS) break;
        shownItems.push(previewItems[i]);
    }

    for (var j = 0; j < shownItems.length; j++) {
        var label = (j + 1) + ". " + String(shownItems[j].name || "");
        list.add("item", label);
    }
    if (previewItems && previewItems.length > shownItems.length) {
        list.add("item", "…（仅显示前 " + shownItems.length + " 条）");
    }

    function visMark(bool) { return bool ? "👀" : "❌"; }
    function clipText(s) {
        var t = String(s || "");
        if (t.length <= STRUCTURE_TEXT_PREVIEW_MAXLEN) return t;
        return t.substring(0, STRUCTURE_TEXT_PREVIEW_MAXLEN) + "…";
    }
    function formatTextPreview(s) {
        var t = String(s || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, "\\n");
        return "✒️\"" + clipText(t) + "\"";
    }

    function formatExcelLine(fields) {
        if (!fields || !fields.length) return "";
        var parts = [];
        for (var i = 0; i < fields.length; i++) {
            var v = fields[i] !== undefined ? String(fields[i]) : "";
            if (EXCEL_ESCAPE_NEWLINES) v = v.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, "\\n");
            parts.push(v);
        }
        return parts.join("\t");
    }

    function strWidth(s) {
        var t = String(s || "");
        var w = 0;
        for (var i = 0; i < t.length; i++) {
            var code = t.charCodeAt(i);
            w += (code <= 255) ? 1 : 2;
        }
        return w;
    }

    function padRight(s, width) {
        var t = String(s || "");
        var w = strWidth(t);
        if (w >= width) return t;
        var n = width - w;
        var pad = "";
        for (var i = 0; i < n; i++) pad += " ";
        return t + pad;
    }

    function formatExcelTable(headerFields, rowFields) {
        var h = headerFields || [];
        var r = rowFields || [];
        var cols = Math.max(h.length, r.length);
        var widths = [];
        for (var i = 0; i < cols; i++) {
            var hs = i < h.length ? String(h[i] !== undefined ? h[i] : "") : "";
            var rs = i < r.length ? String(r[i] !== undefined ? r[i] : "") : "";
            widths.push(Math.max(strWidth(hs), strWidth(rs)));
        }

        function buildLine(fields) {
            var out = "";
            for (var i = 0; i < cols; i++) {
                var v = i < fields.length ? String(fields[i] !== undefined ? fields[i] : "") : "";
                if (EXCEL_ESCAPE_NEWLINES) v = v.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, "\\n");
                if (i === cols - 1) out += v;
                else out += padRight(v, widths[i] + UI_TABLE_COL_GAP);
            }
            return out;
        }

        return buildLine(h) + "\n" + buildLine(r);
    }

    function formatLayerNameMultiline(layerName) {
        var s = String(layerName || "");
        if (UI_LAYERNAME_SPACE_BEFORE_KEY) s = s.replace(/,([^,]*):/g, ",  $1:");
        if (!UI_LAYERNAME_BREAK_ON_COMMA) return s;
        s = s.replace(/,/g, ",\n");
        return s;
    }

    function applyMonospaceSafe(editText) {
        if (!UI_TABLE_USE_MONO_FONT) return;
        try { editText.graphics.font = ScriptUI.newFont(UI_TABLE_MONO_FONT_NAME, "REGULAR", 12); } catch (e) {}
    }

    function renderStructureWithActions(actionMap) {
        if (!SHOW_STRUCTURE) return "";
        var resultLines = [];

        function isPsStatusRelated(layer) {
            var nameStr = String(layer && layer.name ? layer.name : "");
            var isKey = nameStr.length > 0 && nameStr.charAt(nameStr.length - 1) === ":";
            var isParentKey = false;
            try {
                if (layer.parent && layer.parent.typename === "LayerSet") {
                    var pName = String(layer.parent.name || "");
                    isParentKey = pName.length > 0 && pName.charAt(pName.length - 1) === ":";
                }
            } catch (e) {}
            return isKey || isParentKey;
        }

        function getTypeMark(layer) {
            if (!STRUCTURE_LAYER_TYPE_MARK) return "";
            if (layer.typename === "LayerSet") return "[Group] ";
            try {
                if (layer.kind == LayerKind.TEXT) return "[Text] ";
                return "[Layer] ";
            } catch (e) {
                return "[Layer] ";
            }
        }

        function getIndent(depth) {
            if (!STRUCTURE_INDENT_USE_TAB) {
                var s = "";
                for (var i = 0; i < depth; i++) s += String(STRUCTURE_INDENT_CUSTOM || "-");
                return s;
            }
            var t = "";
            for (var j = 0; j < depth; j++) t += "\t";
            return t;
        }

        function traverse(layers, depth) {
            for (var i = 0; i < layers.length; i++) {
                var layer = layers[i];
                var isVisible = getVisibleSafe(layer);
                var passVisible = STRUCTURE_VISIBLE_ONLY ? isVisible : true;
                var passPs = STRUCTURE_PS_STATUS_ONLY ? isPsStatusRelated(layer) : true;
                if (passVisible && passPs) {
                    var line = "";
                    if (STRUCTURE_SHOW_HIERARCHY && depth > 0) line += getIndent(depth) + " ";
                    if (STRUCTURE_VISIBILITY_MARK) line += (isVisible ? "👀 " : "🙈 ");
                    line += getTypeMark(layer);
                    line += String(layer.name || "");

                    var id = getLayerIdSafe(layer);
                    var act = (id !== null && actionMap && actionMap.hasOwnProperty(String(id))) ? actionMap[String(id)] : null;
                    if (act) {
                        var parts = [];
                        if (act.hasOwnProperty("visible")) {
                            var nv = !!act.visible;
                            if (STRUCTURE_SHOW_NOOP_ACTIONS || nv !== isVisible) parts.push(visMark(isVisible) + "=>" + visMark(nv));
                        }
                        if (act.hasOwnProperty("text")) parts.push(formatTextPreview(act.text));
                        if (parts.length > 0) line += "  " + parts.join("  ");
                    }

                    resultLines.push(line);
                }
                if (layer.typename === "LayerSet") traverse(layer.layers, depth + 1);
            }
        }

        traverse(doc.layers, 0);
        if (resultLines.length === 0) return "没有找到符合条件的图层。";
        return resultLines.join("\n");
    }

    function buildDryRunActionMap(configMap) {
        var actionMap = {};
        var jobs = [];
        for (var key in configMap) {
            if (!configMap.hasOwnProperty(key)) continue;
            var pureKey = trim(String(key));
            var targets = targetsByKey.hasOwnProperty(pureKey) ? targetsByKey[pureKey] : [];
            for (var i = 0; i < targets.length; i++) {
                jobs.push({
                    depth: getLayerDepth(targets[i]),
                    target: targets[i],
                    key: pureKey,
                    value: String(configMap[key])
                });
            }
        }
        jobs.sort(function (a, b) { return b.depth - a.depth; });

        function setAct(layer, patch) {
            var id = getLayerIdSafe(layer);
            if (id === null) return;
            var k = String(id);
            if (!actionMap.hasOwnProperty(k)) actionMap[k] = {};
            for (var p in patch) if (patch.hasOwnProperty(p)) actionMap[k][p] = patch[p];
        }

        function isEmptyValue(valueText) {
            var v = trim(String(valueText));
            var lower = v.toLowerCase();
            return v === "" || containsValue(EMPTY_VALUES, lower);
        }

        function isTrueValue(valueText) {
            return containsValue(TRUE_VALUES, trim(String(valueText)).toLowerCase());
        }

        function isFalseValue(valueText) {
            return containsValue(FALSE_VALUES, trim(String(valueText)).toLowerCase());
        }

        function applyKeyToTargetDryRun(target, keyName, rawValue) {
            var valueText = trim(String(rawValue));
            var empty = isEmptyValue(valueText);
            var tVal = isTrueValue(valueText);
            var fVal = isFalseValue(valueText);

            if (psStatusIsKeyTextLayer(target)) {
                if (empty) {
                    setAct(target, { visible: false });
                } else {
                    setAct(target, { visible: true, text: decodeTextValue(valueText) });
                }
                return;
            }

            if (target.typename !== "LayerSet") return;

            var groupType = psStatusDetectKeyGroupType(target, keyName);

            if (groupType.type === "textContent") {
                if (empty) {
                    for (var a = 0; a < target.layers.length; a++) setAct(target.layers[a], { visible: false });
                    return;
                }

                for (var b = 0; b < groupType.keyNameTextLayers.length; b++) {
                    setAct(groupType.keyNameTextLayers[b], { visible: true, text: decodeTextValue(valueText) });
                }
                for (var c = 0; c < target.layers.length; c++) {
                    var child = target.layers[c];
                    if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
                    setAct(child, { visible: true });
                }
                return;
            }

            if (groupType.type === "toggle") {
                if (tVal) {
                    for (var d = 0; d < groupType.keyNameChildren.length; d++) setAct(groupType.keyNameChildren[d], { visible: true });
                    return;
                }
                if (fVal || empty) {
                    for (var e = 0; e < groupType.keyNameChildren.length; e++) setAct(groupType.keyNameChildren[e], { visible: false });
                    return;
                }
                for (var f = 0; f < groupType.keyNameChildren.length; f++) setAct(groupType.keyNameChildren[f], { visible: true });
                return;
            }

            var desiredValues = splitCommaValues(valueText);
            for (var g = 0; g < target.layers.length; g++) {
                var v = target.layers[g];
                if (psStatusIsKeyGroup(v) || psStatusIsKeyTextLayer(v)) continue;
                if (empty) {
                    setAct(v, { visible: false });
                } else {
                    var show = arrayContains(desiredValues, trim(String(v.name)));
                    setAct(v, { visible: show });
                }
            }
        }

        for (var j = 0; j < jobs.length; j++) {
            applyKeyToTargetDryRun(jobs[j].target, jobs[j].key, jobs[j].value);
        }
        return actionMap;
    }

    function refreshRightForIndex(idx) {
        if (idx < 0 || idx >= shownItems.length) {
            rightExcelRaw.text = "";
            rightLayerName.text = "";
            rightStructure.text = "";
            return;
        }
        var item = shownItems[idx];
        var actionMap = buildDryRunActionMap(item.configMap || {});
        rightExcelRaw.text = formatExcelTable(item.headerFields || [], item.rowFields || []);
        rightLayerName.text = formatLayerNameMultiline(item.layerName || "");
        rightStructure.text = renderStructureWithActions(actionMap);

        try {
            var lc = 1 + (rightLayerName.text.match(/\n/g) ? rightLayerName.text.match(/\n/g).length : 0);
            if (lc < 2) lc = 2;
            if (lc > UI_LAYERNAME_MAX_LINES) lc = UI_LAYERNAME_MAX_LINES;
            rightLayerName.preferredSize = [UI_RIGHT_WIDTH, (lc + UI_LAYERNAME_EXTRA_LINES) * UI_LINE_HEIGHT + 8];
            dlg.layout.layout(true);
            dlg.layout.resize();
        } catch (e) {}
    }

    list.onChange = function () {
        try { refreshRightForIndex(list.selection ? list.selection.index : -1); } catch (e) {}
    };

    if (shownItems.length > 0) {
        list.selection = 0;
        refreshRightForIndex(0);
    }

    applyMonospaceSafe(rightExcelRaw);

    var ok = false;
    btnContinue.onClick = function () { ok = true; dlg.close(1); };
    btnExit.onClick = function () { ok = false; dlg.close(0); };

    dlg.show();
    return ok;
}

function psStatusShowNoSelectionDialog() {
    var dlg = new Window("dialog", "PS_STATUS");
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    dlg.add("statictext", undefined, "未选中任何图层。请选择操作：");

    var btns = dlg.add("group");
    btns.alignment = "right";
    var b1 = btns.add("button", undefined, "更新全部状态层");
    var b4 = btns.add("button", undefined, "新建状态层(当前状态)");
    var b5 = btns.add("button", undefined, "Excel 新建");
    var bc = btns.add("button", undefined, "取消");

    var choice = null;
    b1.onClick = function () { choice = "mode1"; dlg.close(1); };
    b4.onClick = function () { choice = "mode4"; dlg.close(1); };
    b5.onClick = function () { choice = "mode5"; dlg.close(1); };
    bc.onClick = function () { choice = null; dlg.close(0); };

    dlg.show();
    return choice;
}

function psStatusCollectTasksFromSelection(selectedLayers, statusGroups) {
    var tasks = [];
    for (var i = 0; i < selectedLayers.length; i++) {
        var layer = selectedLayers[i];
        var sg = psStatusFindAncestorStatusGroup(layer, statusGroups);
        if (!sg) continue;

        if (layer.typename === "LayerSet") {
            var under = [];
            psStatusFindStatusLayersRecursive(layer, under);
            tasks = tasks.concat(under);
        } else {
            tasks.push(layer);
        }
    }
    return uniqueLayers(tasks);
}

function psStatusFindAncestorStatusGroup(layer, statusGroups) {
    var cur = layer;
    while (cur && cur.parent && cur.parent.typename !== "Document") {
        cur = cur.parent;
        if (psStatusIsStatusGroup(cur)) {
            for (var i = 0; i < statusGroups.length; i++) {
                if (statusGroups[i] === cur) return cur;
            }
        }
    }
    return null;
}

function psStatusCollectSelectedTextLayers(selectedLayers) {
    var res = [];
    for (var i = 0; i < selectedLayers.length; i++) {
        var l = selectedLayers[i];
        if (isTextLayer(l)) {
            var t = "";
            try { t = String(l.textItem.contents || ""); } catch (e) { t = ""; }
            if (t && (t.indexOf("\r") >= 0 || t.indexOf("\n") >= 0 || t.indexOf("\t") >= 0)) res.push(l);
        }
    }
    return res;
}

function psStatusCollectAllStatusLayers(statusGroups) {
    var tasks = [];
    for (var i = 0; i < statusGroups.length; i++) {
        psStatusFindStatusLayersRecursive(statusGroups[i], tasks);
    }
    return uniqueLayers(tasks);
}

function psStatusFindStatusLayersRecursive(container, out) {
    if (!container || container.typename !== "LayerSet") return;
    for (var i = 0; i < container.layers.length; i++) {
        var item = container.layers[i];
        if (item.typename === "LayerSet") psStatusFindStatusLayersRecursive(item, out);
        else out.push(item);
    }
}

function psStatusCollectKeyTargetsExcludingStatus(doc) {
    var result = [];
    psStatusCollectKeyTargetsRecursive(doc, result, false);
    return result;
}

function psStatusCollectKeyTargetsRecursive(container, out, inStatus) {
    for (var i = 0; i < container.layers.length; i++) {
        var item = container.layers[i];
        var nextInStatus = inStatus || psStatusIsStatusGroup(item);
        if (nextInStatus) continue;

        if (psStatusIsKeyGroup(item) || psStatusIsKeyTextLayer(item)) out.push(item);
        if (item.typename === "LayerSet") psStatusCollectKeyTargetsRecursive(item, out, nextInStatus);
    }
}

function psStatusIndexTargetsByKey(targets) {
    var map = {};
    for (var i = 0; i < targets.length; i++) {
        var t = targets[i];
        var key = psStatusGetPureKeyName(t.name);
        if (!map.hasOwnProperty(key)) map[key] = [];
        map[key].push(t);
    }
    return map;
}

function psStatusResetAllKeyTargets(allKeyTargets) {
    for (var i = 0; i < allKeyTargets.length; i++) {
        var t = allKeyTargets[i];
        if (psStatusIsKeyTextLayer(t)) {
            trySetVisible(t, false);
            continue;
        }
        if (t.typename === "LayerSet") {
            var keyName = psStatusGetPureKeyName(t.name);
            var groupType = psStatusDetectKeyGroupType(t, keyName);
            if (groupType.type === "textContent") {
                for (var a = 0; a < t.layers.length; a++) trySetVisible(t.layers[a], false);
            } else if (groupType.type === "toggle") {
                for (var b = 0; b < groupType.keyNameChildren.length; b++) trySetVisible(groupType.keyNameChildren[b], false);
            } else {
                for (var c = 0; c < t.layers.length; c++) {
                    var child = t.layers[c];
                    if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
                    trySetVisible(child, false);
                }
            }
        }
    }
}

function psStatusApplyConfigByDepth(targetsByKey, configMap) {
    var jobs = [];
    for (var key in configMap) {
        if (!configMap.hasOwnProperty(key)) continue;
        var pureKey = trim(String(key));
        var targets = targetsByKey.hasOwnProperty(pureKey) ? targetsByKey[pureKey] : [];
        for (var i = 0; i < targets.length; i++) {
            jobs.push({
                depth: getLayerDepth(targets[i]),
                target: targets[i],
                key: pureKey,
                value: String(configMap[key])
            });
        }
    }

    jobs.sort(function (a, b) { return b.depth - a.depth; });

    for (var j = 0; j < jobs.length; j++) {
        psStatusApplyKeyToTarget(jobs[j].target, jobs[j].key, jobs[j].value);
    }
}

function psStatusApplyKeyToTarget(target, keyName, rawValue) {
    var valueText = trim(String(rawValue));
    var lower = valueText.toLowerCase();
    var isEmpty = containsValue(EMPTY_VALUES, lower) || valueText === "";
    var isTrue = containsValue(TRUE_VALUES, lower);
    var isFalse = containsValue(FALSE_VALUES, lower);

    if (psStatusIsKeyTextLayer(target)) {
        if (isEmpty) {
            trySetVisible(target, false);
        } else {
            trySetVisible(target, true);
            psStatusShowLayerAndParentsByPolicy(target);
            setText(target, decodeTextValue(valueText));
        }
        return;
    }

    if (target.typename !== "LayerSet") return;

    var groupType = psStatusDetectKeyGroupType(target, keyName);

    if (groupType.type === "textContent") {
        if (isEmpty) {
            for (var a = 0; a < target.layers.length; a++) trySetVisible(target.layers[a], false);
            return;
        }

        for (var b = 0; b < groupType.keyNameTextLayers.length; b++) {
            var tl = groupType.keyNameTextLayers[b];
            trySetVisible(tl, true);
            psStatusShowLayerAndParentsByPolicy(tl);
            setText(tl, decodeTextValue(valueText));
        }
        for (var c = 0; c < target.layers.length; c++) {
            var child = target.layers[c];
            if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
            trySetVisible(child, true);
            psStatusShowLayerAndParentsByPolicy(child);
        }
        return;
    }

    if (groupType.type === "toggle") {
        if (isTrue) {
            for (var d = 0; d < groupType.keyNameChildren.length; d++) {
                var childOn = groupType.keyNameChildren[d];
                trySetVisible(childOn, true);
                psStatusShowLayerAndParentsByPolicy(childOn);
            }
            return;
        }

        if (isFalse || isEmpty) {
            for (var e = 0; e < groupType.keyNameChildren.length; e++) trySetVisible(groupType.keyNameChildren[e], false);
            return;
        }

        for (var f = 0; f < groupType.keyNameChildren.length; f++) {
            var childDefault = groupType.keyNameChildren[f];
            trySetVisible(childDefault, true);
            psStatusShowLayerAndParentsByPolicy(childDefault);
        }
        return;
    }

    var desiredValues = splitCommaValues(valueText);
    for (var g = 0; g < target.layers.length; g++) {
        var v = target.layers[g];
        if (psStatusIsKeyGroup(v) || psStatusIsKeyTextLayer(v)) continue;
        if (isEmpty) {
            trySetVisible(v, false);
        } else {
            var show = arrayContains(desiredValues, trim(String(v.name)));
            trySetVisible(v, show);
            if (show) psStatusShowLayerAndParentsByPolicy(v);
        }
    }
}

function psStatusDetectKeyGroupType(group, keyName) {
    var keyNameChildren = [];
    var keyNameTextLayers = [];
    for (var i = 0; i < group.layers.length; i++) {
        var child = group.layers[i];
        if (trim(String(child.name)) !== keyName) continue;
        keyNameChildren.push(child);
        if (isTextLayer(child)) keyNameTextLayers.push(child);
    }

    if (keyNameTextLayers.length > 0) return { type: "textContent", keyNameChildren: keyNameChildren, keyNameTextLayers: keyNameTextLayers };
    if (keyNameChildren.length > 0) return { type: "toggle", keyNameChildren: keyNameChildren };
    return { type: "select" };
}

function psStatusValidateConfig(configMap, targetsByKey) {
    var issueDetails = [];
    var issueCount = 0;
    var issueCountMap = {};
    var targetIds = [];

    function getTypeLabelForTarget(target, pureKey) {
        try {
            if (psStatusIsKeyTextLayer(target)) return "keyTextLayer";
            if (target.typename !== "LayerSet") return "unknown";
            var gt = psStatusDetectKeyGroupType(target, pureKey);
            return gt && gt.type ? String(gt.type) : "unknown";
        } catch (e) {
            return "unknown";
        }
    }

    function addIssue(kind, pureKey, target, message, extra) {
        var kindText = String(kind || "");
        var detail = {
            kind: kindText,
            key: String(pureKey || ""),
            message: String(message || ""),
            target: target || null,
            targetId: getLayerIdSafe(target),
            targetPath: target ? psStatusGetLayerPath(target) : "",
            targetType: target ? getTypeLabelForTarget(target, pureKey) : "unknown"
        };
        if (extra) {
            for (var kk in extra) {
                if (extra.hasOwnProperty(kk)) detail[kk] = extra[kk];
            }
        }
        issueDetails.push(detail);
        if (kindText !== "mixed_key_types") {
            var countKey = String(detail.targetPath || "__GLOBAL__") + "\n" + String(detail.message || "");
            if (!issueCountMap.hasOwnProperty(countKey)) {
                issueCountMap[countKey] = true;
                issueCount++;
            }
        }
    }

    for (var key in configMap) {
        if (!configMap.hasOwnProperty(key)) continue;

        var pureKey = trim(String(key));
        if (pureKey.indexOf("__") === 0) continue;
        var valueText = trim(String(configMap[key]));
        var lower = valueText.toLowerCase();
        var isEmpty = containsValue(EMPTY_VALUES, lower) || valueText === "";
        var isTrue = containsValue(TRUE_VALUES, lower);
        var isFalse = containsValue(FALSE_VALUES, lower);

        if (!targetsByKey.hasOwnProperty(pureKey) || targetsByKey[pureKey].length === 0) {
            addIssue("missing_key", pureKey, null, "找不到 key：" + pureKey);
            continue;
        }

        var targets = targetsByKey[pureKey];
        var seenTypeMap = {};
        var typeDetails = [];
        for (var i = 0; i < targets.length; i++) {
            var t = targets[i];
            if (psStatusIsKeyTextLayer(t)) continue;
            if (t.typename !== "LayerSet") continue;

            var groupType = psStatusDetectKeyGroupType(t, pureKey);
            if (!seenTypeMap.hasOwnProperty(groupType.type)) {
                seenTypeMap[groupType.type] = true;
                typeDetails.push({ target: t, type: groupType.type });
            }
            if (groupType.type === "toggle") {
                if (isEmpty || isTrue || isFalse) continue;
                addIssue("invalid_toggle_value", pureKey, t, "key(" + pureKey + ") [type=toggle] 的 value 不是 true/false/空值：" + valueText, { valueText: valueText });
                targetIds.push(getLayerIdSafe(t));
                continue;
            }

            if (groupType.type === "select") {
                if (isEmpty) continue;
                var desired = splitCommaValues(valueText);
                var foundIssue = false;
                for (var k = 0; k < desired.length; k++) {
                    var vv = desired[k];
                    if (!vv) continue;
                    if (!psStatusGroupHasValueChild(t, vv)) {
                        addIssue("missing_select_value", pureKey, t, "key(" + pureKey + ") [type=" + getTypeLabelForTarget(t, pureKey) + "] 找不到 value：" + vv, { valueText: valueText, missingValue: vv });
                        foundIssue = true;
                    }
                }
                if (foundIssue) targetIds.push(getLayerIdSafe(t));
                continue;
            }
        }

        if (typeDetails.length > 1) {
            for (var td = 0; td < typeDetails.length; td++) {
                addIssue("mixed_key_types", pureKey, typeDetails[td].target, "key(" + pureKey + ") 在不同目标位置存在混合类型。", {
                    mixedTypes: psStatusCollectMixedTypeNames(typeDetails),
                    mixedType: typeDetails[td].type
                });
            }
        }
    }
    return { issues: psStatusFormatValidationIssues(issueDetails), issueDetails: issueDetails, issueCount: issueCount, targetIds: targetIds };
}

function psStatusCollectMixedTypeNames(typeDetails) {
    var out = [];
    var seen = {};
    for (var i = 0; i < typeDetails.length; i++) {
        var t = String(typeDetails[i] && typeDetails[i].type ? typeDetails[i].type : "unknown");
        if (seen.hasOwnProperty(t)) continue;
        seen[t] = true;
        out.push(t);
    }
    return out;
}

function psStatusFormatValidationIssues(issueDetails) {
    var lines = [];
    var mixedGroups = {};
    var targetGroups = {};
    var targetOrder = [];
    var mixedOrder = [];

    for (var i = 0; i < issueDetails.length; i++) {
        var it = issueDetails[i];
        if (!it) continue;

        if (it.kind === "mixed_key_types") {
            var mixKey = String(it.key || "");
            if (!mixedGroups.hasOwnProperty(mixKey)) {
                mixedGroups[mixKey] = { key: mixKey, mixedTypes: it.mixedTypes || [], targets: [] };
                mixedOrder.push(mixKey);
            }
            mixedGroups[mixKey].targets.push({ path: String(it.targetPath || ""), type: String(it.mixedType || it.targetType || "unknown") });
            continue;
        }

        var groupKey = it.targetPath ? String(it.targetPath) : "__GLOBAL__";
        if (!targetGroups.hasOwnProperty(groupKey)) {
            targetGroups[groupKey] = { targetPath: String(it.targetPath || ""), targetType: String(it.targetType || "unknown"), messages: [], messageMap: {} };
            targetOrder.push(groupKey);
        }
        if (!targetGroups[groupKey].messageMap.hasOwnProperty(it.message)) {
            targetGroups[groupKey].messageMap[it.message] = true;
            targetGroups[groupKey].messages.push(String(it.message || ""));
        }
    }

    for (var j = 0; j < mixedOrder.length; j++) {
        var mk = mixedOrder[j];
        var mg = mixedGroups[mk];
        if (!mg) continue;
        lines.push("key(" + mg.key + ") 在不同目标位置存在混合类型：" + mg.mixedTypes.join(", "));
        var maxTypeWidth = 0;
        for (var aa = 0; aa < mg.targets.length; aa++) {
            var typeLen = String(mg.targets[aa].type || "").length;
            if (typeLen > maxTypeWidth) maxTypeWidth = typeLen;
        }
        for (var a = 0; a < mg.targets.length; a++) {
            var typeText = String(mg.targets[a].type || "");
            var padCount = maxTypeWidth - typeText.length;
            var typePad = "";
            for (var ap = 0; ap < padCount; ap++) typePad += " ";
            lines.push("- [" + typeText + "]" + typePad + " file://" + mg.targets[a].path);
        }
        lines.push("");
    }

    for (var k = 0; k < targetOrder.length; k++) {
        var tk = targetOrder[k];
        var tg = targetGroups[tk];
        if (!tg) continue;
        if (tg.targetPath) {
            lines.push("file://" + tg.targetPath);
            lines.push("[type=" + tg.targetType + "]");
        } else {
            lines.push("全局配置");
        }
        for (var m = 0; m < tg.messages.length; m++) {
            lines.push("- " + tg.messages[m]);
        }
        if (k < targetOrder.length - 1) lines.push("");
    }

    return lines;
}

function psStatusGroupHasValueChild(group, valueName) {
    for (var i = 0; i < group.layers.length; i++) {
        var child = group.layers[i];
        if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
        if (trim(String(child.name)) === valueName) return true;
    }
    return false;
}

function psStatusFillStatusWithFlatten(mainDoc, targetLayer, statusGroups) {
    var tempDoc = null;
    var oldLayerName = String(targetLayer.name || "");
    var oldParent = targetLayer.parent;

    var savedVisibility = [];
    for (var i = 0; i < statusGroups.length; i++) {
        var sg = statusGroups[i];
        savedVisibility.push({ layer: sg, visible: getVisibleSafe(sg) });
        trySetVisible(sg, false);
    }

    try {
        tempDoc = mainDoc.duplicate();
        app.activeDocument = tempDoc;
        if (tempDoc.layers.length > 1) tempDoc.mergeVisibleLayers();
        var flattenedLayer = tempDoc.activeLayer;
        var layerInMain = flattenedLayer.duplicate(mainDoc, ElementPlacement.PLACEATBEGINNING);
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = mainDoc;

        layerInMain.move(targetLayer, ElementPlacement.PLACEBEFORE);
        layerInMain.name = oldLayerName;
        try { targetLayer.remove(); } catch (e) {}

        for (var r = 0; r < savedVisibility.length; r++) trySetVisible(savedVisibility[r].layer, savedVisibility[r].visible);
        return layerInMain;
    } catch (e2) {
        try { if (tempDoc) tempDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e3) {}
        try { app.activeDocument = mainDoc; } catch (e4) {}
        for (var r2 = 0; r2 < savedVisibility.length; r2++) trySetVisible(savedVisibility[r2].layer, savedVisibility[r2].visible);
        return null;
    } finally {
        try { app.activeDocument = mainDoc; } catch (e5) {}
    }
}

function psStatusBuildConfigFromCurrentState(allKeyTargets) {
    var configMap = {};
    var issues = [];
    var byKey = psStatusIndexTargetsByKey(allKeyTargets);
    for (var key in byKey) {
        if (!byKey.hasOwnProperty(key)) continue;
        var targets = byKey[key];
        var values = [];
        for (var i = 0; i < targets.length; i++) {
            values.push(psStatusGetValueFromTarget(targets[i], key));
        }
        var v0 = values.length > 0 ? values[0] : "";
        for (var j = 1; j < values.length; j++) {
            if (values[j] !== v0) {
                issues.push("同 key(" + key + ") 在不同位置的当前状态不一致。");
                break;
            }
        }
        configMap[key] = v0;
    }
    return { configMap: configMap, issues: uniqueStrings(issues) };
}

function psStatusGetValueFromTarget(target, keyName) {
    if (psStatusIsKeyTextLayer(target)) {
        if (!getVisibleSafe(target)) return "";
        var t = "";
        try { t = String(target.textItem.contents || ""); } catch (e) { t = ""; }
        return encodeTextValue(trim(t));
    }

    if (target.typename === "LayerSet") {
        var gt = psStatusDetectKeyGroupType(target, keyName);
        if (gt.type === "textContent") {
            var anyVisible = false;
            for (var i = 0; i < gt.keyNameTextLayers.length; i++) {
                if (getVisibleSafe(gt.keyNameTextLayers[i])) { anyVisible = true; break; }
            }
            if (!anyVisible) return "";
            var t2 = "";
            try { t2 = String(gt.keyNameTextLayers[0].textItem.contents || ""); } catch (e2) { t2 = ""; }
            return encodeTextValue(trim(t2));
        }
        if (gt.type === "toggle") {
            for (var j = 0; j < gt.keyNameChildren.length; j++) {
                if (getVisibleSafe(gt.keyNameChildren[j])) return "true";
            }
            return "false";
        }

        var visibles = [];
        for (var k = 0; k < target.layers.length; k++) {
            var child = target.layers[k];
            if (psStatusIsKeyGroup(child) || psStatusIsKeyTextLayer(child)) continue;
            if (getVisibleSafe(child)) visibles.push(trim(String(child.name)));
        }
        return visibles.length > 0 ? visibles.join(",") : "";
    }

    return "";
}

function psStatusFindExistingStatusLayerByConfig(statusGroup, normalizedConfig) {
    var layers = [];
    psStatusFindStatusLayersRecursive(statusGroup, layers);
    for (var i = 0; i < layers.length; i++) {
        var m = psStatusParseStatusLayerConfig(layers[i].name);
        if (!m) continue;
        if (psStatusNormalizeConfigMap(m) === normalizedConfig) return layers[i];
    }
    return null;
}

function psStatusNormalizeConfigMap(map) {
    var keys = [];
    for (var k in map) if (map.hasOwnProperty(k)) keys.push(String(k));
    keys.sort();
    var parts = [];
    for (var i = 0; i < keys.length; i++) {
        var kk = keys[i];
        parts.push(trim(kk) + ":" + trim(String(map[kk])));
    }
    return parts.join(",");
}

function psStatusSerializeConfigMap(map, keyOrder) {
    var parts = [];
    if (keyOrder && keyOrder.length) {
        for (var i = 0; i < keyOrder.length; i++) {
            var k = trim(String(keyOrder[i]));
            if (!k) continue;
            if (!map.hasOwnProperty(k)) continue;
            parts.push(k + ":" + String(map[k]));
        }
    } else {
        for (var key in map) {
            if (!map.hasOwnProperty(key)) continue;
            parts.push(trim(String(key)) + ":" + String(map[key]));
        }
    }
    return parts.join(",");
}

function psStatusParseStatusLayerConfig(layerName) {
    var match = String(layerName || "").match(/\{([\s\S]*)\}/);
    if (!match) return null;
    var content = match[1];
    return psStatusParseKeyValueContent(content);
}

function psStatusParseKeyValueContent(content) {
    var tokens = String(content || "").split(",");
    var map = {};
    var currentKey = null;
    for (var i = 0; i < tokens.length; i++) {
        var token = trim(tokens[i]);
        if (!token) continue;
        var idx = token.indexOf(":");
        if (idx >= 0) {
            var k = trim(token.substring(0, idx));
            var v = trim(token.substring(idx + 1));
            currentKey = k;
            map[k] = v;
        } else {
            if (currentKey) map[currentKey] = String(map[currentKey]) + "," + token;
        }
    }
    return map;
}

function psStatusNormalizeMultiLineStatusName(line, index) {
    var s = String(line || "");
    if (s.indexOf("{") >= 0 && s.indexOf("}") >= 0) return s;
    if (s.indexOf(":") >= 0) return "status" + index + "={" + s + "}";
    return "status" + index + "={" + s + "}";
}

function parseTSV(text) {
    try {
        var input = String(text || "");
        input = input
            .replace(/\r\n/g, "\n")
            .replace(/\r/g, "\n")
            .replace(/\u2028/g, "\n")
            .replace(/\u2029/g, "\n")
            .replace(/\u000b/g, "\n")
            .replace(/\u000c/g, "\n");

        function isQuoteChar(ch) {
            return ch === "\"" || ch === "“" || ch === "”";
        }

        var rows = [];
        var row = [];
        var field = "";
        var inQuotes = false;
        var expectedCols = 0;
        var sawHeader = false;

        for (var i = 0; i < input.length; i++) {
            var ch = input.charAt(i);
            if (isQuoteChar(ch)) {
                if (ch === "\"" && inQuotes && i + 1 < input.length && input.charAt(i + 1) === "\"") {
                    field += "\"";
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
                continue;
            }

            if (!inQuotes && ch === "\t") {
                row.push(field);
                field = "";
                continue;
            }

            if (!inQuotes && ch === "\n") {
                if (!sawHeader) {
                    row.push(field);
                    field = "";
                    rows.push(row);
                    expectedCols = row.length;
                    sawHeader = true;
                    row = [];
                    continue;
                }

                if (expectedCols > 0 && row.length < expectedCols - 1) {
                    field += "\n";
                    continue;
                }

                if (expectedCols > 0 && row.length === expectedCols - 1) {
                    var hasTabAhead = false;
                    var j = i + 1;
                    while (j < input.length) {
                        var ch2 = input.charAt(j);
                        if (ch2 === "\n") break;
                        if (ch2 === "\t") { hasTabAhead = true; break; }
                        j++;
                    }
                    if (!hasTabAhead && j > i + 1) {
                        field += "\n";
                        continue;
                    }
                }

                row.push(field);
                field = "";
                var hasAny = false;
                for (var c = 0; c < row.length; c++) {
                    if (String(row[c] || "") !== "") { hasAny = true; break; }
                }
                if (hasAny) rows.push(row);
                row = [];
                continue;
            }

            field += ch;
        }

        if (field !== "" || row.length > 0) {
            row.push(field);
            var hasAny2 = false;
            for (var c2 = 0; c2 < row.length; c2++) {
                if (String(row[c2] || "") !== "") { hasAny2 = true; break; }
            }
            if (hasAny2) rows.push(row);
        }

        lastParseTSVDebugText = "parseTSV: len=" + input.length +
            ", tab=" + (input.indexOf("\t") >= 0) +
            ", nl=" + (input.indexOf("\n") >= 0) +
            ", cols=" + expectedCols +
            ", rows=" + rows.length;

        if (inQuotes) alert("parseTSV 警告：引号未闭合（可能导致换行解析错误）");

        return rows;
    } catch (e) {
        alert("parseTSV 发生错误: " + e.message + "\nLine: " + e.line);
        return [];
    }
}

function splitLines(text) {
    return String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
}

function encodeTextValue(v) {
    return String(v || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, "\\n");
}

function decodeTextValue(v) {
    return String(v || "").replace(/\\n/g, "\r");
}

function psStatusGetTopStatusGroup(doc) {
    var candidates = [];
    for (var i = 0; i < doc.layerSets.length; i++) {
        var ls = doc.layerSets[i];
        if (psStatusIsStatusGroup(ls)) candidates.push(ls);
    }
    if (candidates.length === 1) return candidates[0];
    return null;
}

function psStatusGetTopStatusGroupOrCreate(doc) {
    var candidates = [];
    for (var i = 0; i < doc.layerSets.length; i++) {
        var ls = doc.layerSets[i];
        if (psStatusIsStatusGroup(ls)) candidates.push(ls);
    }
    if (candidates.length === 1) return candidates[0];
    if (candidates.length > 1) return null;
    var sg = doc.layerSets.add();
    sg.name = "status";
    return sg;
}

function psStatusGetOrCreateStatusGroup(container) {
    try {
        var existing = container.layerSets.getByName("status");
        return existing;
    } catch (e) {
        var sg = container.layerSets.add();
        sg.name = "status";
        return sg;
    }
}

function psStatusFindStatusGroups(doc) {
    var res = [];
    psStatusFindStatusGroupsRecursive(doc, res);
    return res;
}

function psStatusFindStatusGroupsRecursive(container, res) {
    for (var i = 0; i < container.layers.length; i++) {
        var item = container.layers[i];
        if (item.typename === "LayerSet") {
            if (psStatusIsStatusGroup(item)) res.push(item);
            psStatusFindStatusGroupsRecursive(item, res);
        }
    }
}

function psStatusIsStatusGroup(layerSet) {
    try { return layerSet && layerSet.typename === "LayerSet" && trim(String(layerSet.name || "")).toLowerCase() === "status"; } catch (e) { return false; }
}

function psStatusIsKeyGroup(layer) {
    try { return layer && layer.typename === "LayerSet" && /[:]$/.test(String(layer.name || "")); } catch (e) { return false; }
}

function psStatusIsKeyTextLayer(layer) {
    try { return layer && layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT && /[:]$/.test(String(layer.name || "")); } catch (e) { return false; }
}

function isTextLayer(layer) {
    try { return layer && layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT; } catch (e) { return false; }
}

function psStatusGetLayerPath(layer) {
    if (!layer) return "";
    var names = [];
    var cur = layer;
    while (cur) {
        names.push(String(cur.name || ""));
        if (!cur.parent || cur.parent.typename === "Document") break;
        cur = cur.parent;
    }
    names.reverse();
    return names.join(" / ");
}

function psStatusGetPureKeyName(name) {
    return trim(String(name || "").replace(/[:]$/, ""));
}

function showLayerAndParents(layer) {
    if (!layer) return;
    trySetVisible(layer, true);
    var p = layer.parent;
    while (p && p.typename !== "Document") {
        trySetVisible(p, true);
        p = p.parent;
    }
}

function psStatusShowLayerAndParentsByPolicy(layer) {
    if (PS_STATUS_SHOW_PARENTS_POLICY === "keep_as_now") return;
    if (PS_STATUS_SHOW_PARENTS_POLICY === "show_parent_only") {
        var p = layer ? layer.parent : null;
        if (p && p.typename !== "Document") trySetVisible(p, true);
        return;
    }
    showLayerAndParents(layer);
}

function getLayerDepth(layer) {
    var d = 0;
    var cur = layer;
    while (cur && cur.parent && cur.parent.typename !== "Document") {
        d++;
        cur = cur.parent;
    }
    return d;
}

function trySetVisible(layer, bool) {
    try { layer.visible = !!bool; } catch (e) {}
}

function getVisibleSafe(layer) {
    try { return !!layer.visible; } catch (e) { return false; }
}

function setText(textLayer, value) {
    try {
        if (!textLayer || textLayer.typename !== "ArtLayer") return;
        if (textLayer.kind !== LayerKind.TEXT) return;
        if (textLayer.allLocked) return;
        if (!textLayer.textItem) return;
        textLayer.textItem.contents = String(value || "");
    } catch (e) {}
}

function splitCommaValues(valueText) {
    var parts = String(valueText || "").split(",");
    var out = [];
    for (var i = 0; i < parts.length; i++) {
        var s = trim(parts[i]);
        if (s) out.push(s);
    }
    return out;
}

function containsValue(arr, lowerVal) {
    for (var i = 0; i < arr.length; i++) {
        if (String(arr[i]).toLowerCase() === lowerVal) return true;
    }
    return false;
}

function arrayContains(arr, val) {
    for (var i = 0; i < arr.length; i++) {
        if (arr[i] === val) return true;
    }
    return false;
}

function uniqueStrings(arr) {
    var map = {};
    var out = [];
    for (var i = 0; i < arr.length; i++) {
        var s = String(arr[i]);
        if (!map.hasOwnProperty(s)) {
            map[s] = true;
            out.push(s);
        }
    }
    return out;
}

function uniqueLayers(arr) {
    var map = {};
    var out = [];
    for (var i = 0; i < arr.length; i++) {
        var id = getLayerIdSafe(arr[i]);
        if (id === null) continue;
        if (!map.hasOwnProperty(id)) {
            map[id] = true;
            out.push(arr[i]);
        }
    }
    return out;
}

function getLayerIdSafe(layer) {
    try { return layer.id; } catch (e) { return null; }
}

function trim(s) {
    return String(s || "").replace(/^\s+|\s+$/g, "");
}

function getSelectedLayers(doc) {
    var selected = [];
    try {
        var ref = new ActionReference();
        ref.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("targetLayersIDs"));
        ref.putEnumerated(charIDToTypeID("Dcmn"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
        var desc = executeActionGet(ref);

        if (desc.hasKey(stringIDToTypeID("targetLayersIDs"))) {
            var ids = desc.getList(stringIDToTypeID("targetLayersIDs"));
            for (var i = 0; i < ids.count; i++) {
                var id = ids.getReference(i).getIdentifier();
                var l = findLayerByID(doc, id);
                if (l) selected.push(l);
            }
        } else {
            selected.push(doc.activeLayer);
        }
    } catch (e) {
        try { selected.push(doc.activeLayer); } catch (e2) {}
    }
    return uniqueLayers(selected);
}

function findLayerByID(parent, id) {
    for (var i = 0; i < parent.layers.length; i++) {
        var layer = parent.layers[i];
        try { if (layer.id == id) return layer; } catch (e) {}
        if (layer.typename === "LayerSet") {
            var found = findLayerByID(layer, id);
            if (found) return found;
        }
    }
    return null;
}

function selectLayer(doc, layer) {
    try {
        doc.activeLayer = layer;
    } catch (e) {}
}

function selectLayersByIds(doc, ids) {
    if (!doc || !ids || ids.length === 0) return;
    try {
        var desc0 = new ActionDescriptor();
        var ref0 = new ActionReference();
        ref0.putIdentifier(charIDToTypeID("Lyr "), ids[0]);
        desc0.putReference(charIDToTypeID("null"), ref0);
        desc0.putBoolean(charIDToTypeID("MkVs"), false);
        executeAction(charIDToTypeID("slct"), desc0, DialogModes.NO);
        for (var i = 1; i < ids.length; i++) {
            var desc = new ActionDescriptor();
            var ref = new ActionReference();
            ref.putIdentifier(charIDToTypeID("Lyr "), ids[i]);
            desc.putReference(charIDToTypeID("null"), ref);
            desc.putEnumerated(stringIDToTypeID("selectionModifier"), stringIDToTypeID("selectionModifierType"), stringIDToTypeID("addToSelection"));
            desc.putBoolean(charIDToTypeID("MkVs"), false);
            executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
        }
    } catch (e) {
        try { if (ids.length > 0) selectLayer(doc, findLayerByID(doc, ids[0])); } catch (e2) {}
    }
}
