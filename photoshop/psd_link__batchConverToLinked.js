//=@show toplevel
/*
    Batch Convert Embedded Smart Objects to Linked (V4 - Robust Fix)
    优化：重构了内嵌状态判断逻辑，改为“排除法”。
    只要图层没有关联明确的外部本地文件路径，就一律判定为内嵌，彻底解决误跳过问题。
*/

// ==================== 配置项 ====================
var DEBUG = true; // 开启调试模式
var ENABLE_CONTENT_DEDUPE = true;
var LAYER_NAME_ROOT_HINTS = "拷贝,副本,Copy,copy";
var ID_TAG_TEMPLATE = "[[{{id}}]]";
var ID_TAG_ID_SOURCE = "timestamp"; // 支持 "layer.id" 或 "timestamp"
var SKIP_NON_PS_EMBEDDED_SMARTOBJECTS = true;
var SKIP_NON_PS_EMBEDDED_SMARTOBJECTS_WITHOUT_OPEN = true;
var ALLOWED_EMBEDDED_SMARTOBJECT_EXTENSIONS = "psd,psb,psdt";
var ALLOWED_EMBEDDED_IMAGE_EXTENSIONS = "png,jpg,jpeg,tif,tiff,gif,bmp,webp";
var SKIP_LAYER_OR_GROUP_NAME_CONTAINS = "toLink_skip,link.js_skip";
var SKIP_LAYER_OR_GROUP_NAME_REGEX = "";
// ===============================================

function main() {
    if (app.documents.length === 0) {
        alert("请先打开一个包含智能对象的 PSD 文档。");
        return;
    }

    var doc = app.activeDocument;

    try {
        var docPath = doc.path;
    } catch (e) {
        alert("请先保存您的文档，脚本需要根据 PSD 的路径来创建外链资源文件夹。");
        return;
    }

    var outputFolder = new Folder(docPath + "/_LinkedObjects");

    var __savedVisibilityByLayerId = null;

    function getLayerIdForLayerObject(layer) {
        try {
            var s2t = stringIDToTypeID;
            var r = new ActionReference();
            r.putProperty(s2t("property"), s2t("layerID"));
            r.putIndex(s2t("layer"), layer.itemIndex);
            var d = executeActionGet(r);
            return d.getInteger(s2t("layerID"));
        } catch (e) {
            return null;
        }
    }

    function saveAllLayersVisibility() {
        var map = {};
        saveVisibilityRecursive(doc.layers, map);
        __savedVisibilityByLayerId = map;
    }

    function saveVisibilityRecursive(layers, map) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            var id = getLayerIdForLayerObject(layer);
            if (id !== null) map[String(id)] = layer.visible;
            if (layer.typename === "LayerSet") {
                saveVisibilityRecursive(layer.layers, map);
            }
        }
    }

    function restoreAllLayersVisibility() {
        if (!__savedVisibilityByLayerId) return;
        restoreVisibilityRecursive(doc.layers, __savedVisibilityByLayerId);
    }

    function restoreVisibilityRecursive(layers, map) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            var id = getLayerIdForLayerObject(layer);
            var key = (id !== null) ? String(id) : null;
            if (key && map.hasOwnProperty(key)) {
                try { layer.visible = map[key] === true; } catch (e) { }
            }
            if (layer.typename === "LayerSet") {
                restoreVisibilityRecursive(layer.layers, map);
            }
        }
    }

    // ==================== 第一阶段：数据扫描与统计 ====================
    var nameRootHints = parseCsvToArray(LAYER_NAME_ROOT_HINTS);
    var allSmartObjectLayers = [];
    var uniqueSOIDs = {};
    var uniqueCount = 0;
    var sampleLayerBySOID = {};
    var layerInfos = [];
    var repLayerBySOID = {};
    var repNameBySOID = {};
    var repScoreBySOID = {};
    var repNumericIdBySOID = {};
    var skippedLayerLines = [];

    saveAllLayersVisibility();

    scanLayers(doc.layers);
    finalizeRepresentatives();

    if (allSmartObjectLayers.length === 0) {
        restoreAllLayersVisibility();
        alert("【扫描结束】未在当前文档中检测到任何符合条件的内嵌智能对象图层。");
        return;
    }

    var contentDedupeInfo = buildContentDedupeInfo();
    restoreAllLayersVisibility();
    var preflight = showPreflightDialog(contentDedupeInfo);
    if (preflight.mode === "cancel") {
        restoreAllLayersVisibility();
        return;
    }

    // ==================== 第三阶段：实际执行转换 ====================
    if (!outputFolder.exists) {
        outputFolder.create();
    }

    var executionCount = 0;
    var exportedSmartObjects = {};
    var DEDUPE_MODE = preflight.mode;
    var ADD_ID_TAG = preflight.addIdTag;
    var dedupeKeyToId = {};
    var timestampCounter = 1;
    var renameQueue = [];
    var layerPlanById = buildLayerPlanMap();

    doc.suspendHistory("批量转换为外链智能对象", "executeConversion()");

    function executeConversion() {
        var errorLines = [];
        var exportErrorsByKey = {};

        for (var k = 0; k < allSmartObjectLayers.length; k++) {
            var createLayerId = allSmartObjectLayers[k];
            var createPlan = layerPlanById[createLayerId];
            if (!createPlan) continue;
            if (exportedSmartObjects[createPlan.dedupeKey] !== undefined) continue;

            selectLayerById(createLayerId);
            var safeName = sanitizeFileName(createPlan.exportBaseName);
            var targetFile = new File(outputFolder + "/" + safeName + ".psb");
            var fileIndex = 1;
            while (targetFile.exists) {
                targetFile = new File(outputFolder + "/" + safeName + "_" + fileIndex + ".psb");
                fileIndex++;
            }

            try {
                convertToLinked(targetFile);
                if (!targetFile.exists) {
                    throw new Error("转换后未生成文件: " + targetFile.fsName);
                }
                exportedSmartObjects[createPlan.dedupeKey] = targetFile;
            } catch (createErr) {
                exportErrorsByKey[createPlan.dedupeKey] = String(createErr);
                errorLines.push("导出失败: " + createPlan.layerName + " | " + createErr);
                if (DEBUG) $.writeln("无法导出图层 " + createPlan.layerName + ": " + createErr);
            }
        }

        for (var m = 0; m < allSmartObjectLayers.length; m++) {
            var relinkLayerId = allSmartObjectLayers[m];
            var relinkPlan = layerPlanById[relinkLayerId];
            if (!relinkPlan) continue;

            var linkedFile = exportedSmartObjects[relinkPlan.dedupeKey];
            if (!linkedFile || !linkedFile.exists) {
                if (!exportErrorsByKey[relinkPlan.dedupeKey]) {
                    errorLines.push("链接失败: " + relinkPlan.layerName + " | 未找到对应的外链文件");
                }
                continue;
            }

            selectLayerById(relinkLayerId);
            try {
                linkCurrentLayerToExistingFile(linkedFile);
                if (ADD_ID_TAG && relinkPlan.desiredLayerName) {
                    renameQueue.push({ id: relinkLayerId, name: relinkPlan.desiredLayerName });
                }
                executionCount++;
            } catch (relinkErr) {
                errorLines.push("链接失败: " + relinkPlan.layerName + " | " + relinkErr);
                if (DEBUG) $.writeln("无法链接图层 " + relinkPlan.layerName + ": " + relinkErr);
            }
        }

        if (renameQueue.length > 0) {
            try {
                applyRenameQueue(renameQueue);
            } catch (e) {
                if (DEBUG) $.writeln("批量重命名失败: " + e);
            }
        }

        restoreAllLayersVisibility();

        var modeLabel = (DEDUPE_MODE === "content") ? "按内容去重" : "按源ID去重";
        var msg = "【转换完成】\n处理方式: " + modeLabel + "\n实际成功处理了 " + executionCount + " 个图层。\n外链文件已保存在：\n" + outputFolder.fsName;
        if (errorLines.length > 0) {
            msg += "\n\n【失败详情】\n" + errorLines.join("\n");
        }
        alert(msg);
    }

    function buildLayerPlanMap() {
        var map = {};
        for (var i = 0; i < layerInfos.length; i++) {
            var info = layerInfos[i];
            var signature = null;
            if (DEDUPE_MODE === "content" && contentDedupeInfo && contentDedupeInfo.soIDToSignature) {
                signature = contentDedupeInfo.soIDToSignature[info.soID] || null;
            }
            var dedupeKey = (DEDUPE_MODE === "content" && signature) ? signature : info.soID;
            var groupId = getGroupIdForDedupeKey(dedupeKey, info.soID, signature);
            var idTag = ADD_ID_TAG ? buildIdTag(groupId) : "";
            var exportBaseName = getExportBaseName(info.soID, signature);
            if (ADD_ID_TAG) exportBaseName = applyIdTagToName(exportBaseName, idTag);

            map[info.id] = {
                layerId: info.id,
                layerName: info.name,
                soID: info.soID,
                signature: signature,
                dedupeKey: dedupeKey,
                exportBaseName: exportBaseName,
                desiredLayerName: ADD_ID_TAG ? applyIdTagToName(info.name, idTag) : null
            };
        }
        return map;
    }

    // ==================== 工具函数群 ====================
    function linkCurrentLayerToExistingFile(fileObj) {
        var lastErr = null;

        try {
            convertToLinked(fileObj);
        } catch (e1) {
            lastErr = e1;
        }
        if (getActiveLayerLinkPath() !== null) return;

        try {
            relinkSmartObject(fileObj);
        } catch (e2) {
            lastErr = e2;
        }
        if (getActiveLayerLinkPath() !== null) return;

        if (lastErr) throw lastErr;
        throw new Error("未能将当前图层转为外链");
    }



    function scanLayers(layers) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];

            if (layer.typename === "LayerSet") {
                var groupSkip = shouldSkipLayerOrGroupByName(layer.name);
                if (groupSkip.skip) {
                    skippedLayerLines.push("跳过组【" + layer.name + "】及其子层。原因: " + groupSkip.reason);
                    continue;
                }
                scanLayers(layer.layers);
            } else if (layer.kind === LayerKind.SMARTOBJECT) {
                var layerSkip = shouldSkipLayerOrGroupByName(layer.name);
                if (layerSkip.skip) {
                    skippedLayerLines.push("跳过图层【" + layer.name + "】。原因: " + layerSkip.reason);
                    continue;
                }
                doc.activeLayer = layer;

                var checkStatus = checkSmartObjectStatus();

                if (checkStatus.isEmbedded) {
                    var layerId = getActiveLayerNumericID();
                    allSmartObjectLayers.push(layerId);

                    var soID = getSmartObjectID();
                    var isFirst = !uniqueSOIDs[soID];
                    if (isFirst) {
                        uniqueSOIDs[soID] = true;
                        uniqueCount++;
                    }
                    considerRepresentativeForSOID(soID, layerId, layer.name);
                    layerInfos.push({
                        id: layerId,
                        name: layer.name,
                        soID: soID
                    });
                } else {
                    skippedLayerLines.push("跳过图层【" + layer.name + "】。原因: " + checkStatus.reason);
                }
            }
        }
    }

    var __SKIP_NAME_CONTAINS_PARTS = null;
    var __SKIP_NAME_REGEX_COMPILED = null;
    var __SKIP_NAME_REGEX_BAD = false;

    function getSkipNameContainsParts() {
        if (__SKIP_NAME_CONTAINS_PARTS) return __SKIP_NAME_CONTAINS_PARTS;
        __SKIP_NAME_CONTAINS_PARTS = parseCsvToArray(SKIP_LAYER_OR_GROUP_NAME_CONTAINS);
        return __SKIP_NAME_CONTAINS_PARTS;
    }

    function compileSkipNameRegex() {
        if (__SKIP_NAME_REGEX_COMPILED || __SKIP_NAME_REGEX_BAD) return __SKIP_NAME_REGEX_COMPILED;
        var raw = String(SKIP_LAYER_OR_GROUP_NAME_REGEX || "");
        if (!raw) return null;

        try {
            var m = raw.match(/^\/([\s\S]*)\/([gimuy]*)$/);
            if (m) {
                __SKIP_NAME_REGEX_COMPILED = new RegExp(m[1], m[2]);
            } else {
                __SKIP_NAME_REGEX_COMPILED = new RegExp(raw);
            }
        } catch (e) {
            __SKIP_NAME_REGEX_BAD = true;
            if (DEBUG) $.writeln("SKIP_LAYER_OR_GROUP_NAME_REGEX 无效，已忽略: " + raw + " | " + e);
            __SKIP_NAME_REGEX_COMPILED = null;
        }
        return __SKIP_NAME_REGEX_COMPILED;
    }

    function shouldSkipLayerOrGroupByName(name) {
        var n = String(name || "");

        var parts = getSkipNameContainsParts();
        for (var i = 0; i < parts.length; i++) {
            var token = String(parts[i] || "");
            if (!token) continue;
            if (n.indexOf(token) !== -1) {
                return { skip: true, reason: "命中包含规则: " + token };
            }
        }

        var re = compileSkipNameRegex();
        if (re) {
            try { re.lastIndex = 0; } catch (e) { }
        }
        if (re && re.test(n)) {
            return { skip: true, reason: "命中正则规则: " + String(re) };
        }

        return { skip: false, reason: "" };
    }

    // 排除法判断逻辑：彻底修复误判
    function checkSmartObjectStatus() {
        var result = { isEmbedded: true, reason: "默认为内嵌智能对象" };
        try {
            var ref = new ActionReference();
            ref.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("smartObjectMore"));
            ref.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
            var desc = executeActionGet(ref);

            if (desc.hasKey(stringIDToTypeID("smartObjectMore"))) {
                var soDesc = desc.getObjectValue(stringIDToTypeID("smartObjectMore"));

                // 核心改动：检查对象内部是否包含外部链接路径（linkPath）键
                // 如果是真正的外链智能对象，底层必然会包含一个指向本地磁盘或云端文件的路径
                if (soDesc.hasKey(stringIDToTypeID("linkPath"))) {
                    var pathStr = soDesc.getString(stringIDToTypeID("linkPath"));
                    if (pathStr && pathStr.length > 0) {
                        result.isEmbedded = false;
                        result.reason = "检测到已有外链路径: " + pathStr;
                        return result;
                    }
                }

                // 兜底验证 type 值
                if (soDesc.hasKey(stringIDToTypeID("type"))) {
                    var type = soDesc.getInteger(stringIDToTypeID("type"));
                    if (DEBUG) $.writeln("图层: " + doc.activeLayer.name + " | type = " + type);
                    // 仅记录 type，不再用它判断外链/内嵌。
                    // 部分 PSD 中内嵌智能对象也会出现 type=2，导致误判为外链从而被跳过。
                }
            }
        } catch (e) {
            // 如果报错，说明结构不标准，极大可能是内嵌，允许转换
            result.isEmbedded = true;
            result.reason = "获取状态异常，默认放行: " + e.message;
        }

        if (result.isEmbedded && SKIP_NON_PS_EMBEDDED_SMARTOBJECTS && SKIP_NON_PS_EMBEDDED_SMARTOBJECTS_WITHOUT_OPEN) {
            try {
                var guess = getActiveLayerSmartObjectSourceNameGuess();
                if (guess) {
                    var ext = getFileExtensionLower(guess);
                    if (ext && !isAllowedEmbeddedSmartObjectExtension(ext)) {
                        result.isEmbedded = false;
                        result.reason = "跳过非 PS/非图像 内嵌智能对象: " + guess;
                        return result;
                    }
                }
            } catch (e2) { }
        }

        return result;
    }

    var __ALLOWED_EMBEDDED_EXT_MAP = null;

    function isAllowedEmbeddedSmartObjectExtension(extLower) {
        var map = getAllowedEmbeddedSmartObjectExtensionMap();
        return map[String(extLower || "").toLowerCase()] === true;
    }

    function getAllowedEmbeddedSmartObjectExtensionMap() {
        if (__ALLOWED_EMBEDDED_EXT_MAP) return __ALLOWED_EMBEDDED_EXT_MAP;
        var parts = parseCsvToArray(String(ALLOWED_EMBEDDED_SMARTOBJECT_EXTENSIONS || "") + "," + String(ALLOWED_EMBEDDED_IMAGE_EXTENSIONS || ""));
        var map = {};
        for (var i = 0; i < parts.length; i++) {
            var k = String(parts[i] || "").toLowerCase();
            if (k) map[k] = true;
        }
        __ALLOWED_EMBEDDED_EXT_MAP = map;
        return map;
    }

    function getFileExtensionLower(pathStr) {
        var s = String(pathStr || "");
        var m = s.match(/\.([0-9a-zA-Z]+)$/);
        if (!m) return "";
        return String(m[1]).toLowerCase();
    }

    function getActiveLayerSmartObjectSourceNameGuess() {
        try {
            var s2t = stringIDToTypeID;
            var ref = new ActionReference();
            ref.putProperty(charIDToTypeID("Prpr"), s2t("smartObject"));
            ref.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
            var desc = executeActionGet(ref);
            var smDesc = desc.getObjectValue(s2t("smartObject"));
            if (smDesc.hasKey(s2t("fileReference"))) {
                try { return smDesc.getString(s2t("fileReference")); } catch (e1) { }
            }
            if (smDesc.hasKey(s2t("placed"))) {
                try { return smDesc.getString(s2t("placed")); } catch (e2) { }
            }
            return null;
        } catch (e) {
            return null;
        }
    }

    function shouldSkipOpenedSmartObjectDoc(docObj) {
        if (!docObj) return false;
        var ext = getDocumentExtensionLower(docObj);
        if (!ext) return false;
        return !isAllowedEmbeddedSmartObjectExtension(ext);
    }

    function getDocumentExtensionLower(docObj) {
        try {
            if (docObj.fullName && docObj.fullName.fsName) return getFileExtensionLower(docObj.fullName.fsName);
        } catch (e1) { }
        try {
            if (docObj.name) return getFileExtensionLower(docObj.name);
        } catch (e2) { }
        return "";
    }

    function parseCsvToArray(csv) {
        var parts = String(csv || "").split(",");
        var arr = [];
        for (var i = 0; i < parts.length; i++) {
            var s = String(parts[i]).replace(/^\s+|\s+$/g, "");
            if (s.length > 0) arr.push(s);
        }
        return arr;
    }

    function countSubStr(haystack, needle) {
        if (!needle) return 0;
        var h = String(haystack);
        var n = String(needle);
        if (n.length === 0) return 0;
        var count = 0;
        var pos = 0;
        while (true) {
            var idx = h.indexOf(n, pos);
            if (idx === -1) break;
            count++;
            pos = idx + n.length;
        }
        return count;
    }

    function getNameScore(name) {
        var n = String(name || "");
        var lower = n.toLowerCase();
        var hits = 0;
        for (var i = 0; i < nameRootHints.length; i++) {
            var m = String(nameRootHints[i]);
            if (m.length === 0) continue;
            var ml = m.toLowerCase();
            hits += countSubStr(lower, ml);
        }
        return { hits: hits, length: n.length, name: n };
    }

    function isBetterNameScore(a, b) {
        if (!b) return true;
        if (a.hits !== b.hits) return a.hits < b.hits;
        if (a.length !== b.length) return a.length < b.length;
        return a.name < b.name;
    }

    function considerRepresentativeForSOID(soID, layerId, layerName) {
        var score = getNameScore(layerName);
        if (isBetterNameScore(score, repScoreBySOID[soID])) {
            repScoreBySOID[soID] = score;
            repLayerBySOID[soID] = layerId;
            repNameBySOID[soID] = layerName;
        }
    }

    function finalizeRepresentatives() {
        for (var soID in repLayerBySOID) {
            if (!repLayerBySOID.hasOwnProperty(soID)) continue;
            sampleLayerBySOID[soID] = repLayerBySOID[soID];
            repNumericIdBySOID[soID] = repLayerBySOID[soID];
        }
        for (var i = 0; i < layerInfos.length; i++) {
            var info = layerInfos[i];
            var repLayer = repLayerBySOID[info.soID];
            info.sourceRootName = repNameBySOID[info.soID] || info.name;
            info.isSourceRoot = (repLayer && info.id === repLayer);
        }
    }

    function buildContentDedupeInfo() {
        if (!ENABLE_CONTENT_DEDUPE) {
            return { enabled: false };
        }

        var soIDToSignature = {};
        var signatureToSOIDs = {};
        var signatureToRepSOID = {};
        var signatureToRepName = {};
        var errors = [];

        for (var soID in sampleLayerBySOID) {
            if (!sampleLayerBySOID.hasOwnProperty(soID)) continue;
            var layerId = sampleLayerBySOID[soID];
            try {
                selectLayerById(layerId);
                if (SKIP_NON_PS_EMBEDDED_SMARTOBJECTS && SKIP_NON_PS_EMBEDDED_SMARTOBJECTS_WITHOUT_OPEN) {
                    var guess = getActiveLayerSmartObjectSourceNameGuess();
                    if (guess) {
                        var ext = getFileExtensionLower(guess);
                        if (ext && !isAllowedEmbeddedSmartObjectExtension(ext)) {
                            soIDToSignature[soID] = null;
                            continue;
                        }
                    }
                }
                var signature = getActiveLayerContentSignature();
                if (!signature) throw new Error("无法生成内容指纹");
                soIDToSignature[soID] = signature;
                if (!signatureToSOIDs[signature]) signatureToSOIDs[signature] = [];
                signatureToSOIDs[signature].push(soID);
            } catch (e) {
                errors.push(doc.activeLayer.name + " | " + e);
                soIDToSignature[soID] = null;
            }
        }

        var signatureCount = 0;
        for (var sig in signatureToSOIDs) {
            if (!signatureToSOIDs.hasOwnProperty(sig)) continue;
            signatureCount++;
            var ids = signatureToSOIDs[sig];
            var repSOID = null;
            var repName = null;
            var repScore = null;
            for (var i = 0; i < ids.length; i++) {
                var id = ids[i];
                var n = repNameBySOID[id];
                if (!n) continue;
                var sc = getNameScore(n);
                if (isBetterNameScore(sc, repScore)) {
                    repScore = sc;
                    repSOID = id;
                    repName = n;
                }
            }
            if (repSOID && repName) {
                signatureToRepSOID[sig] = repSOID;
                signatureToRepName[sig] = repName;
            }
        }

        return {
            enabled: true,
            soIDToSignature: soIDToSignature,
            signatureToSOIDs: signatureToSOIDs,
            signatureToRepSOID: signatureToRepSOID,
            signatureToRepName: signatureToRepName,
            contentUniqueCount: signatureCount,
            errorCount: errors.length,
            errors: errors
        };
    }

    function showPreflightDialog(contentInfo) {
        var sharedLabel = "按源ID去重继续";
        var contentLabel = "按内容去重继续";

        var w = new Window("dialog", "批量转换为外链智能对象");
        w.orientation = "column";
        w.alignChildren = ["fill", "top"];

        var summary = w.add("edittext", undefined, "", { multiline: true, readonly: true });
        summary.preferredSize = [720, 320];

        var lines = [];
        lines.push("预检结果（仅包含内嵌智能对象）");
        lines.push("");
        lines.push("待处理图层总数: " + allSmartObjectLayers.length + " 个");
        lines.push("按源ID去重后的独立数据源: " + uniqueCount + " 个");
        if (contentInfo && contentInfo.enabled) {
            lines.push("按内容去重后的独立数据源: " + contentInfo.contentUniqueCount + " 个");
            if (contentInfo.errorCount && contentInfo.errorCount > 0) {
                lines.push("");
                lines.push("注意: 有 " + contentInfo.errorCount + " 个数据源无法计算内容指纹，将自动按“源ID”处理。");
            }
        } else {
            lines.push("按内容去重: 未启用");
        }
        if (skippedLayerLines.length > 0) {
            lines.push("");
            lines.push("跳过图层: " + skippedLayerLines.length + " 个");
            lines.push(skippedLayerLines.join("\n"));
        }
        lines.push("");
        lines.push("导出位置:");
        lines.push(outputFolder.fsName);
        lines.push("");
        lines.push("图层清单:");
        lines.push(buildReportLines(contentInfo).join("\n"));

        summary.text = lines.join("\n");

        var btnGroup = w.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = ["fill", "top"];

        var addIdCb = btnGroup.add("checkbox", undefined, "增加 Id 标记");
        addIdCb.value = true;
        var spacer = btnGroup.add("group");
        spacer.alignment = ["fill", "fill"];

        var cancelBtn = btnGroup.add("button", undefined, "取消", { name: "cancel" });
        var exportTempBtn = btnGroup.add("button", undefined, "保存临时文件");
        var sharedBtn = btnGroup.add("button", undefined, sharedLabel, {});
        var contentBtn = btnGroup.add("button", undefined, contentLabel, { name: "ok" });

        if (!(contentInfo && contentInfo.enabled)) {
            contentBtn.enabled = false;
        }

        var choice = "cancel";
        cancelBtn.onClick = function () { choice = "cancel"; w.close(); };
        contentBtn.onClick = function () { choice = "content"; w.close(); };
        sharedBtn.onClick = function () { choice = "shared"; w.close(); };
        exportTempBtn.onClick = function () {
            try {
                var folder = exportAllTempContents();
                alert("已保存临时文件到：\n" + folder.fsName);
            } catch (e) {
                alert("保存临时文件失败：\n" + e);
            }
        };

        w.show();
        return { mode: choice, addIdTag: addIdCb.value };
    }

    function getExportBaseName(soID, signature) {
        if (DEDUPE_MODE === "content" && signature && contentDedupeInfo && contentDedupeInfo.signatureToRepName) {
            var rep = contentDedupeInfo.signatureToRepName[signature];
            if (rep) return rep;
        }
        return repNameBySOID[soID] || doc.activeLayer.name;
    }

    function parseIdTagTemplate(tpl) {
        var t = String(tpl || "");
        var marker = "{{id}}";
        var idx = t.indexOf(marker);
        if (idx === -1) return { prefix: t, postfix: "", hasMarker: false };
        return { prefix: t.substring(0, idx), postfix: t.substring(idx + marker.length), hasMarker: true };
    }

    function buildIdTag(idStr) {
        var p = parseIdTagTemplate(ID_TAG_TEMPLATE);
        var id = String(idStr);
        return p.prefix + id + p.postfix;
    }

    function applyIdTagToName(name, idTag) {
        var n = String(name || "");
        var tag = String(idTag || "");
        if (tag.length === 0) return n;
        var p = parseIdTagTemplate(ID_TAG_TEMPLATE);
        if (p.prefix.length > 0 || p.postfix.length > 0) {
            var re = new RegExp(escapeRegExp(p.prefix) + ".*" + escapeRegExp(p.postfix) + "$");
            if (re.test(n)) {
                return n.replace(re, tag);
            }
        } else {
            if (n.indexOf(tag) !== -1) return n;
        }
        return n + tag;
    }

    function escapeRegExp(s) {
        return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    function getGroupIdForDedupeKey(dedupeKey, soID, signature) {
        if (dedupeKeyToId[dedupeKey]) return dedupeKeyToId[dedupeKey];
        var mode = String(ID_TAG_ID_SOURCE || "");
        var idValue = null;

        if (mode === "timestamp") {
            idValue = String(new Date().getTime()) + "_" + String(timestampCounter++);
        } else {
            var repNum = null;
            if (DEDUPE_MODE === "content" && signature && contentDedupeInfo && contentDedupeInfo.signatureToRepSOID) {
                var repSOID = contentDedupeInfo.signatureToRepSOID[signature];
                if (repSOID && repNumericIdBySOID[repSOID] !== undefined) repNum = repNumericIdBySOID[repSOID];
            }
            if (repNum === null && repNumericIdBySOID[soID] !== undefined) repNum = repNumericIdBySOID[soID];
            if (repNum === null) repNum = String(new Date().getTime()) + "_" + String(timestampCounter++);
            idValue = String(repNum);
        }

        dedupeKeyToId[dedupeKey] = idValue;
        return idValue;
    }

    function getLayerNumericID(layer) {
        var prev = doc.activeLayer;
        try {
            doc.activeLayer = layer;
            return getActiveLayerNumericID();
        } catch (e) {
            return null;
        } finally {
            try { doc.activeLayer = prev; } catch (e2) { }
        }
    }

    function getActiveLayerNumericID() {
        var s2t = stringIDToTypeID;
        var r = new ActionReference();
        r.putProperty(s2t("property"), s2t("layerID"));
        r.putEnumerated(s2t("layer"), s2t("ordinal"), s2t("targetEnum"));
        var d = executeActionGet(r);
        return d.getInteger(s2t("layerID"));
    }

    function selectLayerById(layerId) {
        var s2t = stringIDToTypeID;
        var ref = new ActionReference();
        ref.putIdentifier(s2t("layer"), layerId);
        var desc = new ActionDescriptor();
        desc.putReference(s2t("null"), ref);
        executeAction(s2t("select"), desc, DialogModes.NO);
    }

    function buildReportLines(contentInfo) {
        var contentRepSOIDBySOID = {};
        var contentRepNameBySOID = {};
        if (contentInfo && contentInfo.enabled && contentInfo.signatureToSOIDs) {
            for (var sig in contentInfo.signatureToSOIDs) {
                if (!contentInfo.signatureToSOIDs.hasOwnProperty(sig)) continue;
                var ids = contentInfo.signatureToSOIDs[sig];
                if (!ids || ids.length < 2) continue;
                var repSOID = (contentInfo.signatureToRepSOID && contentInfo.signatureToRepSOID[sig]) ? contentInfo.signatureToRepSOID[sig] : null;
                var repName = (contentInfo.signatureToRepName && contentInfo.signatureToRepName[sig]) ? contentInfo.signatureToRepName[sig] : null;
                if (!repSOID || !repName) continue;
                for (var j = 0; j < ids.length; j++) {
                    contentRepSOIDBySOID[ids[j]] = repSOID;
                    contentRepNameBySOID[ids[j]] = repName;
                }
            }
        }

        var lines = [];
        for (var i = 0; i < layerInfos.length; i++) {
            var info = layerInfos[i];
            var tag = info.isSourceRoot ? "[数据源]" : "[共享数据源]";
            var noteParts = [];
            if (!info.isSourceRoot && info.sourceRootName) {
                noteParts.push("与 " + info.sourceRootName + " 共享数据源");
            }
            if (contentInfo && contentInfo.enabled) {
                var repName = contentRepNameBySOID[info.soID];
                var repSOID = contentRepSOIDBySOID[info.soID];
                if (repName && repSOID && repSOID !== info.soID) {
                    noteParts.push("内容与 " + repName + " 相同");
                }
            }

            var line = "  • " + tag + " " + info.name;
            if (noteParts.length > 0) {
                line += " (" + noteParts.join("；") + ")";
            }
            lines.push(line);
        }
        return lines;
    }

    function sanitizeFileName(name) {
        return String(name).replace(/[:\/\\\*\?\"\<\>\|\r\n]/g, "_");
    }

    function makeUniqueFile(folder, baseName, ext) {
        var fileObj = new File(folder.fsName + "/" + baseName + "." + ext);
        var index = 1;
        while (fileObj.exists) {
            fileObj = new File(folder.fsName + "/" + baseName + "_" + index + "." + ext);
            index++;
        }
        return fileObj;
    }

    function exportAllTempContents() {
        var baseFolder = outputFolder;
        if (!baseFolder.exists) baseFolder.create();
        var exportFolder = new Folder(baseFolder.fsName + "/_TempSmartObjectContents");
        if (!exportFolder.exists) exportFolder.create();

        for (var i = 0; i < allSmartObjectLayers.length; i++) {
            var layerId = allSmartObjectLayers[i];
            selectLayerById(layerId);
            var soID = getSmartObjectID();

            var layerName = doc.activeLayer.name;
            var rootName = repNameBySOID[soID] || layerName;
            var baseName = sanitizeFileName(layerName);
            if (rootName && layerName !== rootName) {
                baseName = sanitizeFileName(layerName + " - 与 " + rootName + " 同源");
            }

            saveActiveLayerContentsToFolder(exportFolder, baseName);
        }

        return exportFolder;
    }

    function saveActiveLayerContentsToFolder(folder, baseName) {
        var s2t = stringIDToTypeID;
        if (SKIP_NON_PS_EMBEDDED_SMARTOBJECTS && SKIP_NON_PS_EMBEDDED_SMARTOBJECTS_WITHOUT_OPEN) {
            try {
                var guess = getActiveLayerSmartObjectSourceNameGuess();
                if (guess) {
                    var ext = getFileExtensionLower(guess);
                    if (ext && !isAllowedEmbeddedSmartObjectExtension(ext)) {
                        return;
                    }
                }
            } catch (e0) { }
        }
        executeAction(s2t("placedLayerEditContents"), undefined, DialogModes.NO);
        var soDoc = app.activeDocument;
        try {
            if (SKIP_NON_PS_EMBEDDED_SMARTOBJECTS && shouldSkipOpenedSmartObjectDoc(soDoc)) {
                return;
            }
            var saved = false;
            try {
                var psbFile = makeUniqueFile(folder, baseName, "psb");
                var psbOpts = new LargeDocumentFormatSaveOptions();
                psbOpts.alphaChannels = true;
                psbOpts.layers = true;
                psbOpts.spotColors = true;
                psbOpts.embedColorProfile = true;
                soDoc.saveAs(psbFile, psbOpts, true, Extension.LOWERCASE);
                saved = true;
            } catch (e1) {
                var psdFile = makeUniqueFile(folder, baseName, "psd");
                var psdOpts = new PhotoshopSaveOptions();
                psdOpts.alphaChannels = true;
                psdOpts.layers = true;
                psdOpts.spotColors = true;
                psdOpts.embedColorProfile = true;
                soDoc.saveAs(psdFile, psdOpts, true, Extension.LOWERCASE);
                saved = true;
            }
            if (!saved) {
                throw new Error("无法保存临时文件");
            }
        } finally {
            try { soDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e3) { }
            try { app.activeDocument = doc; } catch (e4) { }
        }
    }

    function getActiveLayerContentSignature() {
        var s2t = stringIDToTypeID;
        executeAction(s2t("placedLayerEditContents"), undefined, DialogModes.NO);
        var soDoc = app.activeDocument;
        try {
            if (SKIP_NON_PS_EMBEDDED_SMARTOBJECTS && shouldSkipOpenedSmartObjectDoc(soDoc)) {
                return null;
            }
            var w = Number(soDoc.width.as("px"));
            var h = Number(soDoc.height.as("px"));
            var mode = String(soDoc.mode);
            var bits = String(soDoc.bitsPerChannel);
            var hist = soDoc.histogram;
            var histCrc = crc32NumberArray(hist);
            var treeCrc = crc32(buildLayerTreeFingerprint(soDoc));
            return w + "x" + h + "_" + mode + "_" + bits + "_" + toHex32(histCrc) + "_" + toHex32(treeCrc);
        } finally {
            try { soDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e1) { }
            try { app.activeDocument = doc; } catch (e2) { }
        }
    }

    function buildLayerTreeFingerprint(soDoc) {
        var parts = [];
        appendLayerContainerFingerprint(parts, soDoc.layers, 0);
        return parts.join("\n");
    }

    function appendLayerContainerFingerprint(parts, layers, depth) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            var indent = "";
            for (var d = 0; d < depth; d++) indent += ">";

            if (layer.typename === "LayerSet") {
                parts.push(indent + "G" + "|" + layerFingerprintCore(layer, true));
                appendLayerContainerFingerprint(parts, layer.layers, depth + 1);
                parts.push(indent + "/G");
            } else {
                parts.push(indent + "L" + "|" + layerFingerprintCore(layer, false));
            }
        }
    }

    function layerFingerprintCore(layer, isGroup) {
        var attrs = [];
        attrs.push("v=" + safeBool(layer.visible));
        attrs.push("op=" + safeNumber(layer.opacity));
        attrs.push("bm=" + safeEnum(layer.blendMode));
        attrs.push("lk=" + safeBool(layer.allLocked));

        if (!isGroup) {
            attrs.push("k=" + safeEnum(layer.kind));
            attrs.push("fo=" + safeNumber(layer.fillOpacity));
            attrs.push("cl=" + safeBool(layer.grouped));
            attrs.push("mt=" + safeBool(layer.isMaskEnabled));
            attrs.push("vn=" + safeBool(layer.isVectorMaskEnabled));
            attrs.push("fx=" + safeBool(layer.layerEffects && layer.layerEffects.visible));

            try {
                if (layer.kind === LayerKind.TEXT && layer.textItem) {
                    var t = String(layer.textItem.contents || "");
                    attrs.push("tc=" + toHex32(crc32(t)));
                }
            } catch (e) { }
        }

        return attrs.join(",");
    }

    function safeBool(v) {
        try { return v ? 1 : 0; } catch (e) { return 0; }
    }

    function safeNumber(v) {
        try {
            var n = Number(v);
            if (isNaN(n)) return 0;
            return n;
        } catch (e) { return 0; }
    }

    function safeEnum(v) {
        try { return String(v); } catch (e) { return ""; }
    }

    var __CRC32_TABLE = null;
    function getCrc32Table() {
        if (__CRC32_TABLE) return __CRC32_TABLE;
        var table = [];
        for (var i = 0; i < 256; i++) {
            var c = i;
            for (var k = 0; k < 8; k++) {
                c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            }
            table[i] = c >>> 0;
        }
        __CRC32_TABLE = table;
        return table;
    }

    function crc32(str) {
        var table = getCrc32Table();
        var crc = 0 ^ (-1);
        for (var i = 0; i < str.length; i++) {
            var b = str.charCodeAt(i) & 0xFF;
            crc = (crc >>> 8) ^ table[(crc ^ b) & 0xFF];
        }
        return (crc ^ (-1)) >>> 0;
    }

    function crc32NumberArray(arr) {
        var table = getCrc32Table();
        var crc = 0 ^ (-1);
        for (var i = 0; i < arr.length; i++) {
            var s = String(arr[i]);
            for (var j = 0; j < s.length; j++) {
                var b = s.charCodeAt(j) & 0xFF;
                crc = (crc >>> 8) ^ table[(crc ^ b) & 0xFF];
            }
            crc = (crc >>> 8) ^ table[(crc ^ 44) & 0xFF];
        }
        return (crc ^ (-1)) >>> 0;
    }

    function toHex32(n) {
        var s = (n >>> 0).toString(16);
        while (s.length < 8) s = "0" + s;
        return s;
    }

    function getSmartObjectID() {
        try {
            var ref = new ActionReference();
            ref.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("smartObjectMore"));
            ref.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
            var desc = executeActionGet(ref);
            var soDesc = desc.getObjectValue(stringIDToTypeID("smartObjectMore"));
            return soDesc.getString(stringIDToTypeID("ID"));
        } catch (e) {
            return "FALLBACK_" + doc.activeLayer.name + "_" + Math.random();
        }
    }

    function convertToLinked(fileObj) {
        var s2t = stringIDToTypeID;

        function doConvert() {
            var desc = new ActionDescriptor();
            var ref = new ActionReference();
            ref.putEnumerated(s2t("layer"), s2t("ordinal"), s2t("targetEnum"));
            desc.putReference(s2t("null"), ref);
            desc.putPath(s2t("using"), fileObj);
            executeAction(s2t("placedLayerConvertToLinked"), desc, DialogModes.NO);
        }

        try {
            doConvert();
        } catch (e) {
            if (!fileObj.exists) {
                executeAction(s2t("newPlacedLayer"), undefined, DialogModes.NO);
                doConvert();
            }
        }
    }

    function getActiveLayerLinkPath() {
        try {
            var s2t = stringIDToTypeID;
            var refSO = new ActionReference();
            refSO.putProperty(s2t("property"), s2t("smartObject"));
            refSO.putEnumerated(s2t("layer"), s2t("ordinal"), s2t("targetEnum"));
            var soDesc = executeActionGet(refSO).getObjectValue(s2t("smartObject"));
            if (soDesc.hasKey(s2t("linked")) && soDesc.getBoolean(s2t("linked")) === true) {
                if (soDesc.hasKey(s2t("link"))) {
                    try {
                        return soDesc.getPath(s2t("link")).fsName;
                    } catch (e1) {
                        return "linked";
                    }
                }
                return "linked";
            }

            return null;
        } catch (e) {
            return null;
        }
    }

    function ensureLinkedToFile(fileObj, allowRecreate) {
        relinkSmartObject(fileObj);
        var p = getActiveLayerLinkPath();
        if (p !== null) return;

        if (!allowRecreate) {
            throw new Error("外链未生效");
        }

        try {
            if (fileObj && fileObj.exists) fileObj.remove();
        } catch (e1) { }

        convertToLinked(fileObj);
        if (!fileObj.exists) {
            throw new Error("无法写入外链文件: " + fileObj.fsName);
        }
        p = getActiveLayerLinkPath();
        if (p === null) {
            throw new Error("外链未生效");
        }
    }

    function applyRenameQueue(items) {
        for (var i = 0; i < items.length; i++) {
            renameLayerById(items[i].id, items[i].name);
        }
    }

    function renameLayerById(layerId, newName) {
        var s2t = stringIDToTypeID;
        var ref = new ActionReference();
        ref.putIdentifier(s2t("layer"), layerId);
        var desc = new ActionDescriptor();
        desc.putReference(s2t("null"), ref);
        var desc2 = new ActionDescriptor();
        desc2.putString(s2t("name"), String(newName));
        desc.putObject(s2t("to"), s2t("layer"), desc2);
        executeAction(s2t("set"), desc, DialogModes.NO);
    }

    function relinkSmartObject(fileObj) {
        if (!fileObj || !fileObj.exists) {
            throw new Error("外链文件不存在: " + (fileObj ? fileObj.fsName : "null"));
        }
        var s2t = stringIDToTypeID;
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putEnumerated(s2t("layer"), s2t("ordinal"), s2t("targetEnum"));
        desc.putReference(s2t("null"), ref);
        desc.putPath(s2t("using"), fileObj);
        executeAction(s2t("placedLayerRelinkToFile"), desc, DialogModes.NO);
    }
}

main();
