#include 'func_itemTree.jsx'
#include 'utils_JSON.jsx'
#include "prototype_Array.jsx"
#include "UI_scriptUI_basicFuncs.jsx"




var emojiDict = {
    lock: "\uD83D\uDD12",    // "🔒",
    hidden: "\uD83D\uDE48",  // "🙈",
    layer: "\uD83D\uDCCB",   // "📋"
    empty: "\u26AA\uFE0F",   // "⚪️"
};


/** * 获取指定 item 的顶层图层
 * @param {PageItem} item 
 * @returns {Layer|null}
 */

function topLayerOf(item) {
    var current = item;
    while (current && current.typename !== "Layer" && current.typename !== "Document") {
        current = current.parent;
    }
    return current && current.typename === "Layer" ? current : null;
}

function getUniqueTopLayerNames(items) {
    if (!items || items.length === 0) return [];
    var seen = {};
    var result = [];
    for (i = 0; i < items.length; i++) {
        var layer = topLayerOf(items[i]);
        if (layer && layer.typename === "Layer") {
            var id = layer.name;
            if (!seen[id]) {
                seen[id] = true;
                result[result.length] = layer.name;
            }
        }
    }

    return result;
}

function itemsTree(items) {
    return getOptimizedItemTree(items, {
        style: "emoji",      // 可选 "emoji" | "ascii" | "basic"
        // indentChar: "ㅤ ",// "＿",    // 使用全角下划线作为缩进
        indentChar: "ㅤ",// "＿",    // 使用全角下划线作为缩进
        showBothIcons: true,  // 显示锁+隐藏两个图标
        showItemType: true,
        showDocName: false,
        showChild: true
    });
}

function checkItemsLockOrHidden(items) {
    if (!Array.isArray) {
        // ES3 兼容：手动判断是否为数组
        if (typeof items.length === "undefined") {
            items = [items];
        } else {
            // 假设是类数组（如 arguments）
            var arr = [];
            for (var i = 0; i < items.length; i++) arr[i] = items[i];
            items = arr;
        }
    }

    var statuses = {
        topLocked: [],
        parentLocked: [],
        itemLocked: [],
        multiLock: [], // 可用于记录同时锁+藏的（可选）

        topHidden: [],
        parentHidden: [],
        itemHidden: [],
        multiHidden: [],

        all: [],

    };

    function isTopLayerLocked(item) {
        var layer = topLayerOf(item);
        return layer && layer.locked;
    }

    function isTopLayerHidden(item) {
        var layer = topLayerOf(item);
        return layer && layer.hidden;
    }

    function hasParentLocked(item) {
        var current = item.parent;
        while (current && current.typename !== "Document") {
            if (current.locked) return true;
            current = current.parent;
        }
        return false;
    }

    function hasParentHidden(item) {
        var current = item.parent;
        while (current && current.typename !== "Document") {
            if (current.hidden) return true;
            current = current.parent;
        }
        return false;
    }

    for (i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;

        var selfLocked = item.locked;
        var selfHidden = item.hidden;
        var topLocked = isTopLayerLocked(item);
        var topHidden = isTopLayerHidden(item);
        var parentLocked = hasParentLocked(item) && !selfLocked;
        var parentHidden = hasParentHidden(item) && !selfHidden;

        if (topLocked) statuses.topLocked.push(item);
        if (parentLocked) statuses.parentLocked.push(item);
        if (selfLocked) statuses.itemLocked.push(item);

        if (topHidden) statuses.topHidden.push(item);
        if (parentHidden) statuses.parentHidden.push(item);
        if (selfHidden) statuses.itemHidden.push(item);

        // 可选：记录同时锁+藏的
        if ((selfLocked || parentLocked || topLocked)) {
            statuses.multiLock.push(item); // 或 multiHidden，按需
        }

        if ((selfLocked || parentLocked || topLocked)) {
            statuses.multiHidden.push(item); // 或 multiHidden ，按需
        }

        if (selfLocked || selfHidden || parentLocked || parentHidden || topLocked || topHidden) {
            statuses.all.push(item);
        }
    }

    return statuses;
}


function showLockHiddenDialogUI(statuses,that) {
    var win = new Window("dialog", "存在隐藏/锁定，请处理");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];

    var headerGroup = win.add('group');
    headerGroup.orientation = "row";
    headerGroup.alignChildren = ["fill", "center"];
    headerGroup.add('statictext', undefined, '检测到以下状态');
    var switchBtn = headerGroup.add('button', undefined, '切换为一览模式');


    // === 上半部分：状态摘要（使用可滚动的只读文本框）===
    var summary = win.add("edittext", undefined, "", {
        multiline: true,
        readonly: true,
        scrolling: true
    });
    summary.size = [600, 600];


    // === 顶层图层：去重显示 ===

    // 辅助函数：安全地拼接字符串（避免 undefined）
    function safeJoin(arr, separator) {
        if (!arr || arr.length === 0) return '';
        return arr.join(separator);
    }

    // 渲染一个状态块的通用函数
    function appendStatusBlock(detailText, config) {
        var items = config.getItems ? config.getItems() : [];
        if (items.length === 0) {
            return detailText;
        }

        var prefixLine = config.emoji + " " + config.label + ": " + items.length + config.unit + "\r";
        var itemLine;

        if (config.formatItem) {
            var formatted = [];
            for (var i = 0; i < items.length; i++) {
                formatted[i] = config.formatItem(items[i]);
            }
            itemLine = safeJoin(formatted, "\n");
        } else {
            itemLine = safeJoin(
                items.map(function (item) {
                    return (config.selfEmoji ? config.selfEmoji : '') + item
                }), "｜");
        }

        return detailText + prefixLine + itemLine + "\r\r";
    }

    function statusesToDetailText(statuses) {
        var detailStatusConfigs = {
            topLocked: {
                emoji: emojiDict.lock,
                label: "顶层图层锁定",
                unit: " 个图层",
                getItems: function () { return getUniqueTopLayerNames(statuses.topLocked); },
                // formatItem: function (name) { return emojiDict.layer + name; }
                selfEmoji: emojiDict.layer,
            },
            topHidden: {
                emoji: emojiDict.hidden,
                label: "顶层图层隐藏",
                unit: " 个图层",
                getItems: function () { return getUniqueTopLayerNames(statuses.topHidden); },
                // formatItem: function (name) { return emojiDict.layer + name; }
                selfEmoji: emojiDict.layer,
            },
            parentLocked: {
                emoji: emojiDict.lock,
                label: "父级锁定",
                unit: " 项",
                getItems: function () { return statuses.parentLocked; },
                formatItem: function (item) { return itemsTree([item]); }
            },
            itemLocked: {
                emoji: emojiDict.lock,
                label: "自身锁定",
                unit: " 项",
                getItems: function () { return statuses.itemLocked; },
                formatItem: function (item) { return itemsTree([item]); }
            },
            parentHidden: {
                emoji: emojiDict.hidden,
                label: "父级隐藏",
                unit: " 项",
                getItems: function () { return statuses.parentHidden; },
                formatItem: function (item) { return itemsTree([item]); }
            },
            itemHidden: {
                emoji: emojiDict.hidden,
                label: "自身隐藏",
                unit: " 项",
                getItems: function () { return statuses.itemHidden; },
                formatItem: function (item) { return itemsTree([item]); }
            }
        };
        var detailText = '';
        // 应用所有配置，构建 detailText
        var detailKeys = ['topLocked', 'topHidden', 'parentLocked', 'itemLocked', 'parentHidden', 'itemHidden'];
        for (var k = 0; k < detailKeys.length; k++) {
            var key = detailKeys[k];
            detailText = appendStatusBlock(detailText, detailStatusConfigs[key]);
        }
        return detailText;
    }


    summary.text = statusesToDetailText(statuses);
    // summary.text = itemsTree(statuses.all);
    // log(summary.text);
    // log("首字符编码: " + summary.text.charCodeAt(0));

    function statusesToJson(statuses) {
        // === 新增：显示 statuses 的 JSON 表示（ES3 兼容）===
        function itemToPlain(item) {
            if (!item || !item.typename) {
                return null;
            }
            return {
                name: item.name ? String(item.name) : "(无名称)",
                typename: String(item.typename)

            };
        }
        // 转换 statuses 为纯数据
        var plainStatuses = {};
        var statusKeys = [
            "topLocked", "parentLocked", "itemLocked",
            "topHidden", "parentHidden", "itemHidden",
            "multiLock", "multiHidden"
        ];

        for (var k = 0; k < statusKeys.length; k++) {
            var key = statusKeys[k];
            var arr = statuses[key];
            if (arr && arr.length) {
                plainStatuses[key] = [];
                for (var i = 0; i < arr.length; i++) {
                    var plain = itemToPlain(arr[i]);
                    if (plain) {
                        plainStatuses[key][plainStatuses[key].length] = plain;
                    }
                }
            } else {
                plainStatuses[key] = [];
            }
        }

        return JSON.stringify(plainStatuses, ' ', 2);
    }
    var showJson = false;
    if (showJson) {

        // 添加 UI
        win.add("statictext", undefined, "原始数据（JSON）：");
        var jsonDetail = win.add("edittext", undefined, jsonText, {
            multiline: true,
            readonly: true,
            scrolling: true
        });
        jsonDetail.size = [400, 150];
        // 使用 JSON.jsx 的 stringify
        var jsonText = statusesToJson(statuses);
        try {
            jsonText = JSON.stringify(plainStatuses, ' ', 2); // 格式化缩进
        } catch (e) {
            jsonText = "JSON.stringify failed: " + e;
        }

    }

    // 确保 switchIdx 在闭包外定义（只初始化一次）
    var switchIdx = 0;
    var btnAndTxt = [{
            btntxt: '切换为一览模式',
            txt: function () { return statusesToDetailText(that.statuses) }
        }, {
            btntxt: '切换为 JSON',
            txt: function () { return itemsTree(that.statuses.all) }
        }, {
            btntxt: '切换为详细内容',
            txt: function () { return statusesToJson(that.statuses) }
        }];

    function refleshContent(){
        that.refreshStatuses();
        var current = btnAndTxt[switchIdx % btnAndTxt.length];
        switchBtn.text = current.btntxt;
        var summaryTxt = current.txt();
        summary.text = summaryTxt.length==0?"已全部解锁/可见":summaryTxt;
        // log("生成前文本: " + summaryTxt);
        // log("首字符编码: " + summaryTxt.charCodeAt(0));
        win.layout.layout(true);
    }

    switchBtn.onClick = function () {
        switchIdx++;  
        refleshContent();
    };



    win.add("statictext", undefined, emojiDict.lock+"锁定 / "+emojiDict.hidden+"隐藏 可能引起脚本错误，请全部解锁/显示: ").textAlign = "left";

    // --- 按钮 ---
    var btnGroup = win.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "fill";                      // ✅ 整个按钮组在窗口中撑满

    btnGroup.alignChildren = ["fill", "center"];      // ✅ 子项（按钮）水平填满、垂直居中

    var _unhideAll = btnGroup.add("button", undefined, "可见全部");
    var _unlockAll = btnGroup.add("button", undefined, "解锁全部");
    var okBtn = btnGroup.add("button", undefined, "下一步");
    var cancelBtn = btnGroup.add("button", undefined, "取消");

    // ✅ 让所有按钮自动等宽撑开
    _unlockAll.alignment = ["fill", "center"];        // ✅ 按钮水平撑满
    _unhideAll.alignment = ["fill", "center"];        // ✅ 按钮水平撑满
    okBtn.alignment = ["fill", "center"];             // ✅ 按钮水平撑满
    cancelBtn.alignment = ["fill", "center"];         // ✅ 按钮水平撑满

    // ✅ （可选）给按钮组一点间距更美观
    btnGroup.spacing = 10;                            // ✅ 按钮间距

    // === 按钮功能 ===
    _unlockAll.onClick = function () {
        var doc = doc || app.activeDocument;
        for (var i = 0; i < doc.layers.length; i++) {
            doc.layers[i].locked = false;
        }
        deepEachPageItemSetProp(doc, "locked", false);
        refleshContent();
        app.redraw();
    };

    _unhideAll.onClick = function () {
        var doc = doc || app.activeDocument;
        for (var i = 0; i < doc.layers.length; i++) {
            doc.layers[i].visible = true;
        }
        deepEachPageItemSetProp(doc, "hidden", false);
        refleshContent();
        app.redraw();
    };

    okBtn.onClick = function () {
        win.close(1);
        applyUserActions({
            unlockAll: unlockAll.value,
            unlockTop: unlockTop.value,
            unlockParent: unlockParent.value,
            unlockItem: unlockItem.value,
            showAll: showAll.value,
            showTop: showTop.value,
            showParent: showParent.value,
            showItem: showItem.value
        }, statuses);
    };

    cancelBtn.onClick = function () {
        win.close(0);
    };





    function applyUserActions(options, statuses) {
        // --- 工具函数 ---
        function findInnermostLayer(item) {
            return item && item.layer; // 直接取 PageItem.layer
        }

        function collectLayerAncestors(layer) {
            var result = [];
            var cur = layer;
            while (cur && cur.typename === "Layer") {
                result.push(cur);
                if (cur.parent && cur.parent.typename === "Document") break;
                cur = cur.parent;
            }
            return result;
        }

        function unlockLayers(layers) {
            for (var i = 0; i < layers.length; i++) {
                if (typeof layers[i].locked !== 'undefined') {
                    layers[i].locked = false;
                }
            }
        }

        function showLayers(layers) {
            for (var i = 0; i < layers.length; i++) {
                if (typeof layers[i].visible !== 'undefined') {
                    layers[i].visible = true;
                }
            }
        }

        function unlockItem(item) {
            if (item && typeof item.locked !== 'undefined') item.locked = false;
        }

        function showItem(item) {
            if (item && typeof item.hidden !== 'undefined') item.hidden = false;
        }

        function unlockParent(item) {
            if (item && item.parent && typeof item.parent.locked !== 'undefined') {
                item.parent.locked = false;
            }
        }

        function showParent(item) {
            var p = item && item.parent;
            if (!p) return;
            if (p.typename === "Layer") {
                if (typeof p.visible !== 'undefined') p.visible = true;
            } else if (typeof p.hidden !== 'undefined') {
                p.hidden = false;
            }
        }

        // --- 批量处理 ---
        function processList(list, fn) {
            if (!list) return;
            for (var i = 0; i < list.length; i++) {
                try { fn(list[i]) } catch (e) {
                    alert(e);
                }
            };
        }

        // --- Top 操作：处理嵌套 Layer 链 ---
        function processTopLocked(list) {
            processList(list, function (item) {
                var layer = findInnermostLayer(item);
                if (layer) {
                    var ancestors = collectLayerAncestors(layer);
                    unlockLayers(ancestors);
                }
            });
        }

        function processTopVisible(list) {
            processList(list, function (item) {
                var layer = findInnermostLayer(item);
                if (layer) {
                    var ancestors = collectLayerAncestors(layer);
                    showLayers(ancestors);
                }
            });
        }


        // --- 执行逻辑 ---
        if (options.unlockAll) {
            processTopLocked(statuses.all);
            processList(statuses.all, unlockItem);

            var doc = doc || app.activeDocument;
            var L = doc.layers;

            for (var i = 0; i < L.length; i++) {
                var l = L[i];
                l.locked = false;
            }
            // for (var i = 0; i < items.length; i++) {
            //     items[i].locked = false;
            // }
            deepEachPageItemSetProp(doc, "locked", false);

        } else {
            if (options.unlockTop) processTopLocked(statuses.topLocked);
            if (options.unlockParent) processList(statuses.parentLocked, unlockParent);
            if (options.unlockItem) processList(statuses.itemLocked, unlockItem);
        }

        if (options.showAll) {
            processTopVisible(statuses.all);
            processList(statuses.all, showItem);

            var doc = doc || app.activeDocument;
            var L = doc.layers;

            for (var i = 0; i < L.length; i++) {
                var l = L[i];
                l.visible = true;
            }

            deepEachPageItemSetProp(doc, "hidden", false);

        } else {
            if (options.showTop) processTopVisible(statuses.topHidden);
            if (options.showParent) processList(statuses.parentHidden, showParent);
            if (options.showItem) processList(statuses.itemHidden, showItem);
        }
        app.redraw();
    }

    var result = win.show();
    if (result !== 1) {
        throw new Error("用户取消操作\n @showLockHiddenDialogUI - UI_checkAskAndUnlockUnhide.js ");
    }

    return result;
}

 





function LockHiddenChecker() {
    this.statuses = null;
    this.config = null;
}


LockHiddenChecker.prototype.check = function (items) {
    this.items = items;
    this.statuses = checkItemsLockOrHidden(items);
    this.refreshStatuses=function(){
        this.statuses = checkItemsLockOrHidden(this.items);
    };
    return this;
};

LockHiddenChecker.prototype.ask = function () {
    this.checkNum = 0; // 统计检查的次数
    if (!this.statuses) throw new Error("请先调用 .check()");
    if (this.statuses.all.length == 0) {
        return true;
    }
    this._continue = showLockHiddenDialogUI(this.statuses,this);

    this.checkNum = 1; // 统计检查的次数

    if (this._continue && this.check(this.items).statuses.all.length != 0 && this.checkNum < 2) {
        if (!confirm("仍然有未处理的锁定/隐藏项目 \n 是否继续？")) {
            this._continue = this.check(this.items).ask();
            this.checkNum = 2; // 统计检查的次数
        }
    }

    return Boolean(this._continue);
};

function testing() {
    var items = app.activeDocument.pageItems; // 你的 item 列表

    var checker = new LockHiddenChecker();
    var checkResult = checker.check(items).ask();

    alert(checkResult);

}

// function log(msg) {
//     try {
//         var now = new Date();
//         var timestamp = now.getFullYear() + "-" +
//             ("0" + (now.getMonth() + 1)).slice(-2) + "-" +
//             ("0" + now.getDate()).slice(-2) + " " +
//             ("0" + now.getHours()).slice(-2) + ":" +
//             ("0" + now.getMinutes()).slice(-2) + ":" +
//             ("0" + now.getSeconds()).slice(-2);

//         var logFile = File(Folder.desktop + "/ai_debug_log.txt");
//         logFile.open("a");
//         logFile.writeln(timestamp + "  " + msg);
//         logFile.close();
//     } catch (e) {
//         throw new Error("日志写入失败: " + e + '\n @log' + " UI_checkAskAndUnlockUnhide.js");
//     }
// }


// testing(); 