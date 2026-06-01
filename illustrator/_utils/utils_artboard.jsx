#include "./prototype_Array.jsx"

// artboard_testing();

function artboard_testing() {
    duplicate_AB_and_CNT_by_ABname('画板 1', 1);
}

function duplicate_AB_and_CNT_by_ABname(artBoardName, copyNum, gap, direction) {

    artBoardName = artBoardName || "ab1";
    copyNum = copyNum || 3;
    gap = gap || 100;
    direction = direction || "right";

    var warn = alert;
    var doc = app.activeDocument;

    var e = false;
    if (getArtboardsByName(doc, artBoardName).length === 0) {
        e = "No artboard named '" + artBoardName + "' found. @ utils_artboard.jsx";
    }
    if (getArtboardsByName(doc, artBoardName).length > 1) {
        e = "Multiple artboards named '" + artBoardName + "' found. @ utils_artboard.jsx";
    }
    if (e) {
        throw new Error(e);
    }

    var contentAB = getArtboardsByName(doc, artBoardName)[0];

    var _contentAB = null;
    for (var idx = 0; idx < copyNum; idx++) {
        _contentAB = duplicate_AB_and_CNT(contentAB, 1, gap, direction)
        _contentAB.name = contentAB.name + (idx + 1);
        contentAB = _contentAB;
    }
    return _contentAB;
}

/**  复制指定画板，并保持其内容不变。

 * @param {Artboard} AB - 要复制的源画板。
 * @param {Number} copyNum - 要复制的次数。
 * @param {Number} gap - 复制画板之间的间距。
 * @param {String} direction - 复制的方向（'right'右, 'down'下, 等）
 * @returns {Artboard} - 新创建的画板
 *
 * @description
*/
function duplicate_AB_and_CNT(AB, copyNum, gap, direction) {
    copyNum = copyNum || 1;
    gap = gap || 400;
    direction = direction || "right";

    var doc = app.activeDocument;

    var e = false;
    if (typeof AB !== "object" || AB.typename !== "Artboard") {
        e = "AB [" + AB + "] is not an artboard. @ utils_artboard.jsx";
    }
    if (e) {
        throw new Error(e);
    }

    var contentAB = AB;

    var _contentAB = null;
    for (var idx = 0; idx < copyNum; idx++) {

        _contentAB = duplicateEmptyArtboard(contentAB, "right", gap);
        duplicateArtboardContent(contentAB, _contentAB);
        _contentAB.name = AB.name + (idx + 1);
        contentAB = _contentAB;
    }
    return _contentAB;
}

function indexOfArtboard(AB) {
    var doc = doc || app.activeDocument;
    if (!doc || !doc.typename || doc.typename != "Document") {
        throw "Invalid document @indexOfArtboard @utils_artboard.js";
    }
    if (!AB || !AB.typename || AB.typename != "Artboard") {
        throw "Invalid artboard @indexOfArtboard @utils_artboard.js";
    }
    for (var i = 0; i < doc.artboards.length; i++) {
        if (doc.artboards[i] == AB) {
            return i;
        }
    }
    throw "Artboard not found on doc " + doc.name + "  @indexOfArtboard @utils_artboard.js";

}

/**
 * 复制所有完全位于指定旧画板范围内的页面对象，
 * 并将复制的对象移动到新画板的位置，保持它们的相对位置不变。
 *
 * @param {Artboard} oldArtboard - 要复制内容的源画板。
 * @param {Artboard} newArtboard - 放置复制内容的目标画板。
 * */

function copyAndPasteInPlaceArtboardContent(oldArtboard, newArtboard) {
    var doc = doc || app.activeDocument;
    var oldRect = oldArtboard.artboardRect;
    var newRect = newArtboard.artboardRect;
    var dx = newRect[0] - oldRect[0];
    var dy = newRect[1] - oldRect[1];

    var prevSelection = doc.selection;

    // 选择原画板所有内容
    doc.selection = null;
    doc
        .artboards
        .setActiveArtboardIndex(indexOfArtboard(oldArtboard));
    app.executeMenuCommand('selectallinartboard');

    if (doc.selection.length == 0) {
        throw new Error('画板' + oldArtboard.name + '内无对象或不可选中，请检查\n@duplicateArtboardContent utils_artboard.jsx');
    }

    // 直接复制整个选择集 doc.artboards.setActiveArtboardIndex(indexOfArtboard(oldArtboard));
    app.executeMenuCommand('copy');

    // 粘贴到新画板

    doc
        .artboards
        .setActiveArtboardIndex(indexOfArtboard(newArtboard));
    app.executeMenuCommand('pasteInPlace');

    doc.selection = prevSelection;
}

/**
 * 分别在不同图层，批量处理指定画板中的所有页面对象。
 *
 * @param {Artboard} artboard - 要处理的画板。
 * @param {Function} fn - 处理页面对象的函数。
 */
function processArtboardsByLayers(artboard1, fn1, artboard2, fn2, fn3) {
    var doc = app.activeDocument;
    var prevActiveLayer = doc.activeLayer;
    var prevActiveAB = doc
        .artboards
        .getActiveArtboardIndex();

    var idx1 = indexOfArtboard(artboard1);
    // 1. 处理 artboard1：按图层分组并执行 fn1
    var groups1 = getArtboardLayerGroups(artboard1);
    for (var i = 0; i < groups1.length; i++) {
        var group = groups1[i];
        doc.activeLayer = group.layer;
        doc
            .artboards
            .setActiveArtboardIndex(idx1);
        if (typeof fn1 === "function") {
            fn1(group.items); // 传入 PageItem[]
        }

        // 2. 执行中间操作 fn2（如 pasteInPlace）
        if (artboard2 && typeof fn2 === "function") {
            var idx2 = indexOfArtboard(artboard2);
            doc
                .artboards
                .setActiveArtboardIndex(idx2);
            fn2(); // 无参数，由用户控制上下文
        }
    }

    // 3. 处理 artboard2：重新获取其内容并执行 fn3
    if (artboard2 && typeof fn3 === "function") {
        var groups2 = getArtboardLayerGroups(artboard2);
        for (var j = 0; j < groups2.length; j++) {
            var group = groups2[j];
            doc.activeLayer = group.layer;
            fn3(group.items); // 传入 PageItem[]
        }
    }

    // 恢复状态
    doc.activeLayer = prevActiveLayer;
    doc
        .artboards
        .setActiveArtboardIndex(prevActiveAB);

}

// 获取画板中所有对象，按图层分组
/**
     * 获取指定画板中的所有页面对象，并按图层分组。
     *
     * @param {Artboard} artboard - 要获取页面对象的画板。
     * @returns {Array} - 一个数组，包含按图层分组的页面对象。
     */
function getArtboardLayerGroups(artboard) {
    var doc = app.activeDocument;
    var allLayers = getAllLayers(doc);
    var result = [];

    for (var i = 0; i < allLayers.length; i++) {
        var layer = allLayers[i];
        var itemsInLayer = getItemsInLayer(layer);
        var itemsInArtboard = [];

        for (var j = 0; j < itemsInLayer.length; j++) {
            var item = itemsInLayer[j];
            if (isItemInArtboard(item, artboard)) {
                itemsInArtboard.push(item);
            }
        }

        if (itemsInArtboard.length > 0) {
            result.push({layer: layer, items: itemsInArtboard});
        }
    }
    return result;
}

/**
     * 获取图层中所有顶层 PageItem（不递归 Group）
     * @param {Layer} layer - The layer to get items from
     * @return {PageItem[]} - An array of PageItems in the layer
     */
function getItemsInLayer(layer) {
    var items = [];
    if (layer && layer.pageItems) {
        for (var i = layer.pageItems.length - 1; i >= 0; i--) {
            items.push(layer.pageItems[i]);
        }
    }
    return items;
}

/**
    * 判断对象是否在画板内（边界相交检测）
    * @param {PageItem} item - The item to check
    * @param {Artboard} artboard - The artboard to check against
    * @return {boolean} - True if the item's bounds intersect with the artboard, false otherwise
    */
function isItemInArtboard(item, artboard) {
    try {
        // 获取对象的几何边界 [left, top, right, bottom]
        var gb = item.geometricBounds;
        var itemLeft = gb[0];
        var itemTop = gb[1];
        var itemRight = gb[2];
        var itemBottom = gb[3];

        // 获取画板边界 [left, top, right, bottom]
        var ab = artboard.artboardRect;
        var abLeft = ab[0];
        var abTop = ab[1];
        var abRight = ab[2];
        var abBottom = ab[3];

        // 检查两个矩形是否相交 不相交的条件：item 在画板左侧、右侧、上方或下方
        var noIntersection = (itemRight < abLeft || // item 完全在画板左边
                itemLeft > abRight || // item 完全在画板右边
                itemBottom > abTop || // item 完全在画板上边（注意：AI y 轴向下）
                itemTop < abBottom // item 完全在画板下边
        );

        return !noIntersection;
    } catch (e) {
        return false;
    }
}

/**
     * 获取文档所有图层（含嵌套）
     * @param {GroupItem} container
     * @param {Array} result
     */
function getAllLayers(container, result) {
    if (!result) 
        result = [];
    var layers = container.layers || (container.typename === "Document"
        ? container.layers
        : []);
    for (var i = layers.length - 1; i >= 0; i--) {
        var lyr = layers[i];
        result.push(lyr);
        if (lyr.layers && lyr.layers.length > 0) {
            getAllLayers(lyr, result);
        }
    }
    return result;
}

function duplicateArtboardContent(oldArtboard, newArtboard) {
    processArtboardsByLayers(oldArtboard, function (items) { // ✅ 接收 items
        var doc = app.activeDocument;
        doc.selection = items; // ✅ 必须先选中
        app.executeMenuCommand('copy');
    }, newArtboard, function () {
        app.executeMenuCommand('pasteInPlace');
    }
    // fn3 可省略
    );
}

/**
 * 在指定方向上复制画板，并保持指定间距
 *
 * @param {Artboard} AB - 要复制的源画板
 * @param {String} direction - 复制的方向（'right'右, 'down'下, 等）
 * @param {Number} gap - 原画板和复制画板之间的间距
 * @returns {Artboard} - 新创建的画板
 *
 * @description
 * 在指定方向上创建一个空白画板的副本。
 * 新画板保持与原画板相同的尺寸，
 * 但按指定的间距距离放置。
 *
 * @example
 * var sourceArtboard = app.activeDocument.artboards[0];
 * duplicateEmptyArtboard(sourceArtboard, "right", 100);
 */

function duplicateEmptyArtboard(AB, direction, gap) {
    var doc = doc || app.activeDocument;
    var LIMIT = { // pt
        LEFT: -7840,
        TOP: 7700,
        RIGHT: 8400,
        BOTTOM: -8530
    };

    var rect = AB.artboardRect;
    var left = rect[0],
        top = rect[1],
        right = rect[2],
        bottom = rect[3];
    var abW = right - left,
        abH = bottom - top;
    var l = left,
        t = top;

    var moveit = function () {
        switch (direction) {
            case "right":
                left += abW + gap;
                if (left + abW > LIMIT.RIGHT) {
                    left = LIMIT.LEFT + (abW + gap);
                    top += (abH - gap);
                }
                break;
            case "left":
                left -= abW + gap;
                if (left < LIMIT.LEFT) {
                    left = LIMIT.RIGHT - (abW + gap);
                    top += (abH - gap);
                }
                break;
            case "up":
            case "top":
                top -= abH - gap;
                if (top > LIMIT.TOP) {
                    top = LIMIT.BOTTOM + (-abH + gap);
                    left += (abW + gap);
                }
                break;
            case "down":
            case "bottom":
                top += abH - gap;
                if (top + abH < LIMIT.BOTTOM) {
                    top = LIMIT.TOP + (abH - gap);
                    left -= (abW + gap);
                }
                break;
        }
    }
    moveit();

    while (isArtboardExist([
        left, top, left + abW,
        top + abH
    ])) {
        moveit();

        if (left < LIMIT.LEFT || left + abW > LIMIT.RIGHT || top < LIMIT.BOTTOM || top + abH > LIMIT.TOP) {
            throw new Error("No available space for new artboard.");
        }
    }

    // 检查部分重叠
    function isArtboardExist(rect) {
        for (var i = 0; i < doc.artboards.length; i++) {
            var abRect = doc.artboards[i].artboardRect;
            // 判断是否有重叠
            if (rect[0] < abRect[2] && rect[2] > abRect[0] && // 水平有重叠
            rect[1] > abRect[3] && rect[3] < abRect[1] // 垂直有重叠
            ) {
                return true;
            }
        }
        return false;
    }

    try {
        return doc
            .artboards
            .add([
                left, top, left + abW,
                top + abH
            ]);
    } catch (e) {
        throw new Error('添加画板失败\n@duplicateEmptyArtboard utils_artboard.jsx\n' + e);
        // alert([e,'\n',abW,abH,'\n',l,t,'\n',left, top,'\n',LIMIT.LEFT,LIMIT.TOP,
        // '\n\n',LIMIT.RIGHT,LIMIT.BOTTOM].join(','));
    }
}

/**
 * 获取所有符合条件的画板。
 * @param {Document} doc - Illustrator 文档对象
 * @param {string|RegExp} name - 画板名称（字符串或正则表达式）
 * @param {boolean} [exactMatch=true] - 是否完全匹配（仅当 name 为字符串时有效）
 * @returns {Artboard[]} - 返回所有匹配的 artboard 数组
 */
function getArtboardsByName(doc, name, exactMatch) {
    var result = [];
    var isReg = (name instanceof RegExp);
    if (typeof exactMatch !== "boolean") 
        exactMatch = true;
    for (var i = 0; i < doc.artboards.length; i++) {
        try {
            var abn = doc.artboards[i].name;
            if (isReg) {
                if (name.test(abn)) 
                    result.push(doc.artboards[i]);
                }
            else if (exactMatch) {
                if (abn === name) 
                    result.push(doc.artboards[i]);
                }
            else {
                if (abn.indexOf(name) !== -1) 
                    result.push(doc.artboards[i]);
                }
            } catch (e) {
            throw new Error("未找到名称为 '" + name + "' 的画板。\n@getArtboardsByName - utils_artboard.jsx\n" + e);
        }
    }
    return result;
}

function normalizeName(n) {
    try {
        return String(n)
            .replace(/[{}]/g, "")
            .replace(/^\s+|\s+$/g, "");
    } catch (e) {
        return "";
    }
}
