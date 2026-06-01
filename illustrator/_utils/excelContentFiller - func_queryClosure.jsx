// 已测试 ✅
// testing();


function getPageItemsFrom_GPT(obj, options) {

    // 判断 item 是否在 roots 下（自身或者任一祖先匹配）
    var isDescendantOf = function (item, roots) {
        if (!item) return false;
        for (var r = 0; r < roots.length; r++) {
            if (!roots[r]) continue;
            if (item === roots[r]) return true;
        }

        var p = item.parent;
        while (p) {
            for (var k = 0; k < roots.length; k++) {
                if (p === roots[k]) return true;
            }
            // 若到 document 则结束
            try {
                if (p.typename && p.typename === "Document") break;
            } catch (e) {             }
            p = p.parent;
        }
        return false;
    }

    // 将 Collection 转为普通数组（并可选对单个 item 做过滤）
    var arrayFrom = function (col, options) {
        var arr = [];
        if (!col) return arr;
        for (var i = 0; i < col.length; i++) {
            var it = col[i];
            if (options) {
                if (options.skipHidden) {
                    try { if (it.hidden) continue; } catch (e) { }
                }
                if (options.skipLocked) {
                    try { if (it.locked) continue; } catch (e) { }
                }
            }
            arr.push(it);
        }
        return arr;
        // ES3 兼容的 "is array-like" 检查（包括真正的 Array 或 Array-like collection）
    }
    var isArrayLike = function (x) {
        if (!x) return false;
        try {
            // 真正的 Array 在 ExtendScript 中也会通过 toString 检测
            var t = Object.prototype.toString.call(x);
            if (t === "[object Array]") return true;
        } catch (e) { }
        // 其次判断是否有 numeric length 属性（Array-like / Collection）
        return (typeof x.length === "number");
    }

    // 备用：在极端情况下没有 id，就回退到 parent-chain 构造 key
    var getReferenceKey = function (item) {
        try {
            var parts = [];
            var cur = item;
            var depth = 0;
            while (cur && depth < 50) {
                var tn = cur.typename || "";
                var nm = "";
                try { nm = (cur.name !== undefined && cur.name !== "") ? cur.name : ""; } catch (e) { }
                parts.push(tn + ":" + nm);
                if (cur.typename === "Document") break;
                cur = cur.parent;
                depth++;
            }
            return parts.join("|");
        } catch (e) {
            return String(item);
        }
    }


    if (!options) options = {};
    var doc = app.activeDocument;
    if (!doc) return [];

    // 如果传入文档或未传入 -> 全部 pageItems
    if (!obj || (obj.typename && obj.typename === "Document")) {
        return arrayFrom(doc.pageItems, options);
    }

    // roots 数组化（ES3 兼容的 isArray 检测）
    var roots = [];
    if (isArrayLike(obj) && !obj.typename) {
        for (var i = 0; i < obj.length; i++) if (obj[i]) roots.push(obj[i]);
    } else {
        roots.push(obj);
    }

    // 遍历 doc.pageItems（权威全集），判断每个 item 是否属于 roots
    var all = doc.pageItems;
    var out = [];
    var seen = {}; // 用 item.id 去重

    for (var j = 0; j < all.length; j++) {
        var it = all[j];

        // 可选过滤：跳过隐藏 / 锁定
        if (options.skipHidden) {
            // TextFrames/Paths 等都有 hidden 属性（容错）
            try { if (it.hidden) continue; } catch (e) { }
        }
        if (options.skipLocked) {
            try { if (it.locked) continue; } catch (e) { }
        }

        if (isDescendantOf(it, roots)) {
            // 用 id 去重（ExtendScript 中 pageItem.id 是稳定的）
            var key = "";
            try { key = String(it.id); } catch (e) { key = getReferenceKey(it); }
            if (!seen[key]) {
                seen[key] = true;
                out.push(it);
            }
        }
    }

    return out;

}


function getPageItemsFrom_qwen(obj) { // qwen new
    var doc = app.activeDocument;
    if (obj === doc || obj === undefined || obj === null) {
        return arrayFrom(doc.pageItems);
    }

    var roots = [];
    // 判断是否为 selection 或 collection（类数组）
    if (obj && typeof obj.length === "number" && !obj.typename) {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i] !== null) roots.push(obj[i]);
        }
    } else {
        // 单个对象
        roots.push(obj);
    }

    var result = [];
    for (var i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];
        if (isDescendantOf(item, roots)) {
            result.push(item);
        }
    }
    return result;
    function isDescendantOf(item, roots) {
        // 检查 item 本身是否在 roots 中
        for (var j = 0; j < roots.length; j++) {
            if (item === roots[j]) return true;
        }
        // 向上遍历 parent
        var p = item.parent;
        while (p && p.typename !== "Document") {
            for (var k = 0; k < roots.length; k++) {
                if (p === roots[k]) return true;
            }
            p = p.parent;
        }
        return false;
    }

    function arrayFrom(col) {
        var arr = [];
        if (!col) return arr;
        for (var i = 0; i < col.length; i++) {
            arr.push(col[i]);
        }
        return arr;
    }
}


/**
 * collectPageItemsByRoots(obj)
 * - 如果 obj 是 Document：直接返回 doc.pageItems 的数组（与 doc.pageItems.length 一致）
 * - 如果 obj 是 selection（Array-like）或单个 Layer/Group/Item：遍历 doc.pageItems，
 *   将属于这些 roots（本身或任一祖先在 roots 中）的 pageItems 收集出来（去重）。
 *
 * 用法：
 *   var items = collectPageItemsByRoots(app.activeDocument); // 全部：等同于 doc.pageItems
 *   var items = collectPageItemsByRoots(app.selection); // 选区（包含其子孙）
 *
 */

function collectPageItemsByRoots(obj, doc) {

    // helper: 将 Illustrator 的 collection 转为普通数组
    var arrayFromCollection = function (col) {
        var arr = [];
        try {
            if (!col) return arr;
            for (var i = 0; i < col.length; i++) arr.push(col[i]);
        } catch (e) { }
        return arr;
    }

    // helper: 生成稳定的唯一键（尽量基于 parent 链与 typename / name / position）
    var getReferenceKey = function (item) {
        try {
            var parts = [];
            var cur = item;
            // 往上到 document 为止，收集 typename + name/index/position 做备份
            var depth = 0;
            while (cur && depth < 50) {
                var t = cur.typename || "";
                var n = "";
                try { n = (cur.name !== undefined && cur.name !== "") ? cur.name : ""; } catch (e) { }
                var pos = "";
                try { pos = (cur.position && cur.position.length) ? cur.position.join(",") : ""; } catch (e) { }
                parts.push(t + ":" + n + "@" + pos);
                if (cur.typename === "Document") break;
                cur = cur.parent;
                depth++;
            }
            return parts.join("|");
        } catch (e) {
            return String(item);
        }
    }


    if (!doc) doc = app.activeDocument;
    var all = arrayFromCollection(doc.pageItems); // authoritative list

    // 形成 roots 数组（把 selection / single / collection 都转成普通数组 of objects）
    var roots = [];
    if (obj === undefined || obj === null) {
        // 默认为文档（全部）
        roots = [doc];
    } else if (obj.typename === "Document") {
        // 明确传入文档 -> 直接返回全部
        return all.slice(0);
    } else if (obj instanceof Array || (obj && typeof obj.length === "number" && !obj.typename)) {
        // selection 或 collection（Array-like）
        for (var i = 0; i < obj.length; i++) {
            if (obj[i]) roots.push(obj[i]);
        }
    } else {
        // 单个对象（Layer, GroupItem, PageItem, etc）
        roots.push(obj);
    }

    // 特殊情况：如果 roots 包含文档对象，则直接返回全部
    for (var r = 0; r < roots.length; r++) {
        if (roots[r] && roots[r].typename === "Document") {
            return all.slice(0);
        }
    }

    // 标记去重
    var seen = {};
    var out = [];

    // 判断 item 是否属于任一 root（item 本身或任一 ancestor === root）
    function isUnderRoots(item, rootsArr) {
        if (!item) return false;
        // 先快速比较 item 是否等于某个 root
        for (var i = 0; i < rootsArr.length; i++) {
            if (item === rootsArr[i]) return true;
        }
        // 向上遍历 parent 链
        var p = item.parent;
        while (p) {
            for (var j = 0; j < rootsArr.length; j++) {
                if (p === rootsArr[j]) return true;
            }
            // 如果到达 Document 或 Layer 顶端可以再上走一次（some parents might be layer/document)
            p = p.parent;
        }
        return false;
    }

    for (var k = 0; k < all.length; k++) {
        var pi = all[k];
        if (isUnderRoots(pi, roots)) {
            // 用 item 内建 index/uuid 不可靠，使用 reference 来去重
            var key = getReferenceKey(pi);
            if (!seen[key]) {
                seen[key] = true;
                out.push(pi);
            }
        }
    }

    return out;

}




/**
 * 获取所有 PageItems（兼容 Document、Layer、GroupItem、selection 等）
 * 并自动去重，确保结果与 doc.pageItems 一致。
 */
function getAllPageItems(obj, recursive) {
    var isCollection = function (obj) {
        try {
            return (obj && typeof obj.length === "number" && !(obj instanceof String) && !obj.typename);
        } catch (e) {
            return false;
        }
    }

    var getUniqueItemKey = function (item) {
        try {
            // 每个 pageItem 有唯一的索引结构：文档名 + 图层路径 + 坐标 + typename
            var name = item.name || "";
            var parentName = (item.parent && item.parent.name) || "";
            var pos = "";
            try {
                pos = item.position ? item.position.join(",") : "";
            } catch (e) { }
            return [item.typename, parentName, name, pos].join("|");
        } catch (e) {
            return String(item);
        }
    }

    if (recursive === undefined) recursive = true;

    var result = [];
    var seen = {}; // 用于去重

    function collect(item) {
        if (!item) return;

        // 数组或 selection
        if (item instanceof Array || (item.length && !item.typename)) {
            for (var i = 0; i < item.length; i++) collect(item[i]);
            return;
        }

        // Document
        if (item.typename === "Document") {
            for (var i = 0; i < item.pageItems.length; i++) collect(item.pageItems[i]);
            return;
        }

        // Layer
        if (item.typename === "Layer") {
            for (var i = 0; i < item.pageItems.length; i++) collect(item.pageItems[i]);
            if (recursive) {
                for (var j = 0; j < item.layers.length; j++) collect(item.layers[j]);
            }
            return;
        }

        // GroupItem
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) collect(item.pageItems[i]);
            if (recursive) {
                for (var k = 0; k < item.groupItems.length; k++) collect(item.groupItems[k]);
            }
            return;
        }

        // PageItems 集合
        if (isCollection(item)) {
            for (var n = 0; n < item.length; n++) collect(item[n]);
            return;
        }

        // PageItem
        if (item.typename && item.parent) {
            var uid = getUniqueItemKey(item);
            if (!seen[uid]) {
                seen[uid] = true;
                result.push(item);
            }
            return;
        }
    }

    collect(obj);
    return result;

}


/**
 * 获取容器（Document 或 Selection）中的所有 PageItem（叶子对象）
 * - 如果是 Document/Layer/Group：直接返回 .pageItems（扁平、无组）
 * - 如果是 Selection：递归展开其中的 GroupItem，返回所有叶子
 * @param {Object} container - app.activeDocument 或 app.selection
 * @returns {Array} PageItem 数组（不含 GroupItem）
 */
function getPageItems(container) {
    if (!container) return [];

    // 情况1: 是 Selection（类数组，可能包含 GroupItem）
    if (container === app.selection || (typeof container.length === 'number' && !container.pageItems)) {
        var result = [];
        for (var i = 0; i < container.length; i++) {
            var item = container[i];
            if (item.typename === "GroupItem") {
                // 递归获取组内所有 PageItem
                result = result.concat(getPageItems(item));
            } else if (item.hasOwnProperty("typename") && item.typename !== "Layer") {
                result.push(item);
            }
        }
        return result;
    }

    // 情况2: 是 Document / Layer / GroupItem → 直接使用 .pageItems（已是扁平叶子列表）
    if (container.pageItems) {
        var arr = [];
        var pis = container.pageItems;
        for (var j = 0; j < pis.length; j++) {
            arr.push(pis[j]);
        }
        return arr;
    }

    return [];
}

function getAllLeafPageItems(container) {
    if (!container) return [];

    var items = [];

    if (container === app.selection || (typeof container.length === 'number' && container.length >= 0)) {
        // Selection：需要递归，因为 selection 只包含直接选中的项（可能包含组）
        for (var i = 0; i < container.length; i++) {
            var item = container[i];
            if (item.typename === "GroupItem") {
                // 递归获取组内所有叶子项
                items = items.concat(getAllLeafPageItems(item));
            } else {
                items.push(item);
            }
        }
    } else if (container.pageItems !== undefined) {
        // Document / Layer / Group：直接使用 .pageItems（它已经是扁平的！）
        var pis = container.pageItems;
        for (var j = 0; j < pis.length; j++) {
            // 注意：这里 pis[j] 永远不会是 GroupItem！
            // 因为 .pageItems 只包含“叶子”对象（PathItem, TextFrame, etc.）
            items.push(pis[j]);
        }
    }

    return items;
}

function extractParametersAdvanced(text, config) {

    var AdvancedConfig = {
        BRACE_OPEN: '{',
        BRACE_CLOSE: '}',
        // TEMPLATE_DELIMITERS: ['`', '"', "'"]  // 支持多种模板分隔符
        TEMPLATE_DELIMITERS: ['`']  // 支持多种模板分隔符
    };


    var parameters = {};
    var result = [];
    var currentConfig = config || AdvancedConfig;

    if (!text || typeof text !== 'string') {
        return result;
    }

    var pos = 0;
    var length = text.length;

    while (pos < length) {
        var _char = text.charAt(pos);

        // 处理 {} 参数
        if (_char === currentConfig.BRACE_OPEN) {
            var braceEnd = findMatchingBraceAdvanced(text, pos, currentConfig);
            if (braceEnd !== -1) {
                var param = text.substring(pos, braceEnd + 1);
                parameters[param] = true;
                pos = braceEnd + 1;
                continue;
            }
        }

        // 处理模板参数（支持多种分隔符）
        if (currentConfig.TEMPLATE_DELIMITERS) {
            var isTemplateChar = false;
            var templateChar = null;

            for (var i = 0; i < currentConfig.TEMPLATE_DELIMITERS.length; i++) {
                if (_char === currentConfig.TEMPLATE_DELIMITERS[i]) {
                    isTemplateChar = true;
                    templateChar = currentConfig.TEMPLATE_DELIMITERS[i];
                    break;
                }
            }

            if (isTemplateChar) {
                var templateEnd = text.indexOf(templateChar, pos + 1);
                if (templateEnd !== -1) {
                    var templateParam = text.substring(pos, templateEnd + 1);
                    parameters[templateParam] = true;
                    pos = templateEnd + 1;
                    continue;
                }
            }
        }

        pos++;
    }

    // 将对象键转换为数组
    for (var key in parameters) {
        if (parameters.hasOwnProperty(key)) {
            result.push(key);
        }
    }

    return result;
}

function findMatchingBraceAdvanced(text, start, config) {
    var currentConfig = config || AdvancedConfig;
    var depth = 1;
    var pos = start + 1;
    var length = text.length;
    var inTemplate = false;
    var currentTemplateChar = null;

    while (pos < length && depth > 0) {
        var _char = text.charAt(pos);

        // 处理模板字符串的开始和结束
        if (currentConfig.TEMPLATE_DELIMITERS && !inTemplate) {
            for (var i = 0; i < currentConfig.TEMPLATE_DELIMITERS.length; i++) {
                if (_char === currentConfig.TEMPLATE_DELIMITERS[i]) {
                    inTemplate = true;
                    currentTemplateChar = currentConfig.TEMPLATE_DELIMITERS[i];
                    break;
                }
            }
        } else if (inTemplate && _char === currentTemplateChar) {
            inTemplate = false;
            currentTemplateChar = null;
        }

        // 如果在模板字符串内，跳过所有特殊字符处理
        if (inTemplate) {
            pos++;
            continue;
        }

        // 处理花括号嵌套
        if (_char === currentConfig.BRACE_OPEN) {
            depth++;
        } else if (_char === currentConfig.BRACE_CLOSE) {
            depth--;
        }

        pos++;
    }

    if (depth === 0) {
        return pos - 1;
    }

    return -1;
}
function getItemsByName(name) {
    var doc = doc || app.activeDocument;
    var result = [];
    for (var i = doc.pageItems.length - 1; i >= 0; i--) {
        var item = doc.pageItems[i];
        if (item.name === name) {
            result.push(item);
        }
    }

    if (result.length > 0) {
        return result;
    }

    throw new Error("未找到名称为 '" + name + "' 的图像。");
}

/**
* 通用函数：从 Document 或 Selection 中提取所有 PageItem（叶子对象）
* - 如果传入 Document/Layer/Group：返回其 .pageItems（扁平列表，不含 Group）
* - 如果传入 Selection：递归展开其中的 Group，返回所有叶子 PageItem
* @param {Object} container - 可以是 app.activeDocument 或 app.selection
* @returns {Array} PageItem 数组（只包含路径、文本、图像等，不含 GroupItem）
*/
function extractPageItems(container) {
    if (!container) return [];

    // 判断是否为 Selection：selection 没有 .pageItems 属性，但有 .length
    if (container === app.selection || (typeof container.length === 'number' && container.length >= 0 && !container.pageItems)) {
        var result = [];
        for (var i = 0; i < container.length; i++) {
            var item = container[i];
            if (item.typename === "GroupItem") {
                // 递归展开组
                result = result.concat(extractPageItems(item));
            } else if (item.typename && item.typename !== "Layer") {
                // 只收集真正的 PageItem（排除 Layer 等）
                result.push(item);
            }
        }
        return result;
    }

    // 否则认为是 Document / Layer / GroupItem：直接使用 .pageItems（已是扁平叶子列表）
    if (container.pageItems) {
        var arr = [];
        var pis = container.pageItems;
        for (var j = 0; j < pis.length; j++) {
            arr.push(pis[j]);
        }
        return arr;
    }

    return [];
}





/** * 查询文档中包含参数的页面项并按参数名分组
 * @param {Object} container - Illustrator文档对象
 * @param {Object} config - 配置对象，用于定义参数解析规则
 * @param {string} config.BRACE_OPEN - 参数开始符号，默认为'{'
 * @param {string} config.BRACE_CLOSE - 参数结束符号，默认为'}'
 * @param {string[]} config.TEMPLATE_DELIMITERS - 模板分隔符数组，默认为['`']
 * @returns {Object} 按参数名分组的对象，包含keys属性存储所有参数名
 */
function queryGet(container, config, selectFn) {
    container = container || app.activeDocument;
    config = config || {
        BRACE_OPEN: '{',
        BRACE_CLOSE: '}',
        // TEMPLATE_DELIMITERS: ['`', '"', "'"]  // 支持多种模板分隔符
        TEMPLATE_DELIMITERS: ['`']  // 支持多种模板分隔符
    };

    selectFn = selectFn || function (target) {
        var items = [];

        // === 1. 如果是 selection ===
        if (target === app.selection) {
            var sel = app.selection;
            for (var i = 0; i < sel.length; i++) {
                items = items.concat(getAllPageItems(sel[i]));
            }
            return items;
        }

        // === 2. 如果是 Document ===
        if (target.typename == "Document") {
            var docItems = target.pageItems;
            for (var i = 0; i < docItems.length; i++) {
                items.push(docItems[i]);
            }
            return items;
        }

        // === 3. 如果有 pageItems 属性（如 Layer、GroupItem 等） ===
        if (target.pageItems) {
            var subItems = target.pageItems;
            for (var i = 0; i < subItems.length; i++) {
                items.push(subItems[i]);
            }
            return items;
        }

        // === 4. 如果就是一个 PageItem ===
        if (target.typename && target.parent) {
            items.push(target);
            return items;
        }

        return items;
    };


    var ParametersInName;
    var ParametersGroup = {};
    ParametersGroup.keys = [];
    ParametersGroup.items = [];

    var k = '';
    // 遍历文档中的所有页面项，提取参数并按参数名分组
    var items = selectFn(container); 

    ParametersGroup.selectItems = items;

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        ParametersInName = extractParametersAdvanced(item.name, config);
        if (ParametersInName.length > 0) {
            for (var j = 0; j < ParametersInName.length; j++) {
                k = ParametersInName[j];
                var _k = k;
                if (k == 'keys') {
                    k = '_keys';
                }
                if (k == 'items') {
                    k = '_items';
                }

                if (k != _k) {
                    if (!confirm('参数名' + _k + '冲突，是否修改为 ' + _k + '并继续？')) {
                        throw new Error('参数名' + _k + '冲突，请修改参数名！');
                    }
                }

                if (typeof ParametersGroup[k] === 'undefined') {
                    ParametersGroup[k] = [];
                    ParametersGroup.keys.push(k);
                }
                ParametersGroup[k].push(item);
                ParametersGroup.items.push(item);

            }
        }
    }
    return ParametersGroup;
}

/**
 * 在指定 item 的下一层 children（pageItems）中查找名称匹配的对象
 * @param {Object} item - Layer 或 GroupItem
 * @param {String|RegExp} matcher - 要匹配的字符串或正则表达式
 * @returns {Array} 匹配的子项数组 [child(ren)...]
 */
function findChildrenByName(item, matcher) {
    var children = [];
    if (!item || !item.pageItems) return children;

    var isRegex = matcher instanceof RegExp;
    var pageItems = item.pageItems;

    for (var i = 0; i < pageItems.length; i++) {
        var child = pageItems[i];
        var name = child.name || "";

        if (isRegex) {
            if (matcher.test(name)) {
                children.push(child);
            }
        } else if (typeof matcher === "string") {
            if (name.indexOf(matcher) !== -1) {
                children.push(child);
            }
        }
    }

    return children;
}



// alert(n);
function testing() {
    var doc = doc || app.activeDocument;
    var confirm = confirm || function (msg) {
        return alert(msg);
    };

    var sel = doc.selection;
    // sel = sel.length==0?doc:sel; 

    var n = '\n';

    var content1 = ['',
        'doc.pageitems.length :' + doc.pageItems.length,
        'doc.selection.length :' + doc.selection.length,
        '\n',
        'queryGet(sel,false,false): ' + queryGet(sel, false, false).items.length,
        'queryGet(doc,false,false): ' + queryGet(doc, false, false).items.length,
        '\n',
        'queryGet(sel,false,getPageItemsFrom_GPT): ' + queryGet(sel, false, getPageItemsFrom_GPT).items.length,
        'queryGet(doc,false,getPageItemsFrom_GPT): ' + queryGet(doc, false, getPageItemsFrom_GPT).items.length,
        '\n',
        'queryGet(sel,false,getPageItemsFrom_qwen): ' + queryGet(sel, false, getPageItemsFrom_qwen).items.length,
        'queryGet(doc,false,getPageItemsFrom_qwen): ' + queryGet(doc, false, getPageItemsFrom_qwen).items.length,
        '\n',
        'queryGet(sel,false,collectPageItemsByRoots): ' + queryGet(sel, false, collectPageItemsByRoots).items.length,
        'queryGet(doc,false,collectPageItemsByRoots): ' + queryGet(doc, false, collectPageItemsByRoots).items.length,
        '\n', '\n',
        'queryGet(sel,false,getAllPageItems): ' + queryGet(sel, false, getAllPageItems).items.length,
        'queryGet(doc,false,getAllPageItems): ' + queryGet(doc, false, getAllPageItems).items.length,
        'queryGet(sel,false,getPageItems): ' + queryGet(sel, false, getPageItems).items.length,
        'queryGet(doc,false,getPageItems): ' + queryGet(doc, false, getPageItems).items.length,
        'queryGet(sel,false,getAllLeafPageItems): ' + queryGet(sel, false, getAllLeafPageItems).items.length,
        'queryGet(doc,false,getAllLeafPageItems): ' + queryGet(doc, false, getAllLeafPageItems).items.length
    ].join('\n');

    confirm(content1);


    confirm(['\n',

        'doc.pageitems.length :' + doc.pageItems.length,
        'doc.selection.length :' + doc.selection.length,
        '\n',
        'queryGet(sel,false,false).selectItems : ' + queryGet(sel, false, false).selectItems.length,
        'queryGet(doc,false,false).selectItems : ' + queryGet(doc, false, false).selectItems.length,
        '\n',
        'queryGet(sel,false,getPageItemsFrom_GPT).selectItems : ' + queryGet(sel, false, getPageItemsFrom_GPT).selectItems.length,
        'queryGet(doc,false,getPageItemsFrom_GPT).selectItems : ' + queryGet(doc, false, getPageItemsFrom_GPT).selectItems.length,
        '\n',
        'queryGet(sel,false,getPageItemsFrom_qwen).selectItems : ' + queryGet(sel, false, getPageItemsFrom_qwen).selectItems.length,
        'queryGet(doc,false,getPageItemsFrom_qwen).selectItems : ' + queryGet(doc, false, getPageItemsFrom_qwen).selectItems.length,
        '\n',
        'queryGet(sel,false,collectPageItemsByRoots).selectItems : ' + queryGet(sel, false, collectPageItemsByRoots).selectItems.length,
        'queryGet(doc,false,collectPageItemsByRoots).selectItems : ' + queryGet(doc, false, collectPageItemsByRoots).selectItems.length,
        '\n', '\n',
        'queryGet(sel,false,getAllPageItems).selectItems: ' + queryGet(sel, false, getAllPageItems).selectItems.length,
        'queryGet(doc,false,getAllPageItems).selectItems: ' + queryGet(doc, false, getAllPageItems).selectItems.length,
        'queryGet(sel,false,getPageItems).selectItems: ' + queryGet(sel, false, getPageItems).selectItems.length,
        'queryGet(doc,false,getPageItems).selectItems: ' + queryGet(doc, false, getPageItems).selectItems.length,
        'queryGet(sel,false,getAllLeafPageItems).selectItems: ' + queryGet(sel, false, getAllLeafPageItems).selectItems.length,
        'queryGet(doc,false,getAllLeafPageItems).selectItems: ' + queryGet(doc, false, getAllLeafPageItems).selectItems.length

    ].join('\n'));




    var ParametersGroup = queryGet(sel);

    n = '';
    n += ParametersGroup.keys.join('\n') + '\n\n\n';
    for (var i = 0; i < ParametersGroup.keys.length; i++) {
        k = ParametersGroup.keys[i];
        n += k + ' : \n' + ParametersGroup[k].length + '个\n' + ParametersGroup[k].join('\n') + '\n\n\n';
    }

    confirm('\n' + ParametersGroup.keys.join('\n'));
    confirm('\n' + n);

}
