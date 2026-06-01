function ungroup(group) {
    if (!(group instanceof GroupItem)) return;

    var parent = group.parent;        // 获取父对象
    var items = group.pageItems;      // 获取组内所有子项
    var len = items.length;

    // 因为移出组会改变索引，所以从后向前遍历
    for (var i = len - 1; i >= 0; i--) {
        var item = items[i];
        item.moveBefore(group);       // 将子项移到组的前面（同级）
    }

    group.remove();                   // 删除空组
}

function group() {
    var items = [];

    // --- 参数展开 ---
    if (arguments.length === 1) {
        var arg = arguments[0];
        if (!arg) return null;

        if (arg instanceof Array) {
            items = arg;
        } else if (arg.hasOwnProperty("length") && !(arg instanceof String)) {
            // selection 或 PageItems 集合
            for (var i = 0; i < arg.length; i++) items.push(arg[i]);
        } else {
            items = [arg];
        }
    } else {
        // 多个独立参数
        for (var j = 0; j < arguments.length; j++) {
            items.push(arguments[j]);
        }
    }

    if (items.length < 1) return null;

    // --- 获取第一个有效 parent ---
    var parent = null;
    for (var k = 0; k < items.length; k++) {
        if (items[k] && items[k].parent) {
            parent = items[k].parent;
            break;
        }
    }
    if (!parent) return null;

    // --- 创建组 ---
    var groupItem = parent.groupItems.add();

    // --- 将 items 移入组 ---
    for (var n = items.length - 1; n >= 0; n--) {
        var it = items[n];
        if (it && it.parent) {
            try {
                it.moveToBeginning(groupItem);
            } catch (e) {
                // 跳过 locked/hidden 无法移动的对象
                throw new Error("无法移动对象"+ it.name+'\n group() utils_group.jsx');
            }
        }
    }

    return groupItem;
}

/**
 * 根据 items 的顶层 Layer 分组
 * @param {...any} args - 可能是多个 item，或一个包含 items 的数组
 * @returns {Object} { obj, str }
 */
function groupItemsByTopLayer() {
    /**
     * 获取 item 所在的顶层 Layer
     * @param {PageItem} item
     * @returns {Layer|null}
     */
    function getTopLayer(item) {
        var current = item;
        while (current && current.typename !== "Layer" && current.typename !== "Document") {
            current = current.parent;
        }
        if (current && current.typename === "Layer") 
            return current;
        return null;
    }

    var args = arguments;
    var items = [];

    // === 统一输入格式 ===
    if (args.length === 1 && args[0]instanceof Array) {
        items = args[0];
    } else {
        for (var i = 0; i < args.length; i++) {
            items.push(args[i]);
        }
    }

    var resultObj = {};

    // === 遍历所有 item ===
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) 
            continue;
        
        var topLayer = getTopLayer(item);
        if (!topLayer) 
            continue;
        
        var layerName = topLayer.name || "(无名称图层)";
        var itemName = item.name || "(未命名对象)";

        if (!resultObj[layerName]) 
            resultObj[layerName] = [];
        resultObj[layerName].push(itemName);
    }

    // === 生成字符串输出 ===
    var lines = [];
    for (var key in resultObj) {
        lines.push("_______");
        lines.push(key);
        lines.push("__");
        lines.push(resultObj[key].join(" , "));
        lines.push("_______");
    }

    var resultStr = lines.join("\n");

    return {obj: resultObj, str: resultStr};
}