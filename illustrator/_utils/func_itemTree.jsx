#include './utils_JSON.jsx'
#include "./prototype_Array.jsx"
#include "./UI_func_uiDialog.jsx"

// 已测试 ✅
// testing();


/**
 * 为每个 item 生成唯一 ID 并存储在 note 中
 * 使用 JSON 格式，保留原来的 note 内容
 * @param {PageItem} item 
 * @returns {string} 生成的 uniqueID
 */
function generateUniqueID(item) {
    if (!item) return null;

    // 读取原 note 内容
    var noteData = {};
    if (item.note) {
        try {
            noteData = JSON.parse(item.note);
        } catch (e) {
            // 如果原来不是 JSON，先存入 _originalNote
            noteData = { _originalNote: item.note };
        }
    }

    // 如果已有 uniqueID，就直接返回
    if (noteData.uniqueID) return noteData.uniqueID;

    // 生成新的 uniqueID（这里用时间戳+随机数）
    var uniqueID = 'id_' + new Date().getTime() + '_' + Math.floor(Math.random()*1000000);

    noteData.uniqueID = uniqueID;

    // 写回 note
    item.note = JSON.stringify(noteData);

    return uniqueID;
}

function allItemUniqueIDs() {
    var ids = [];
    var items = app.activeDocument.pageItems;
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var id = item.name+' : '+generateUniqueID(item);
        if (id) ids.push(id);
    }
    return ids;
}

/**
 * 获取 item.note 中存储的 JSON 属性
 * @param {PageItem} item 
 * @param {string} key - 要读取的属性名，比如 'uniqueID'
 * @returns {any} 对应值，如果不存在返回 null
 */
function itemDetail(item, key) {
    if (!item || !item.note) return null;
    try {
        var data = JSON.parse(item.note);
        return data[key] !== undefined ? data[key] : null;
    } catch (e) {
        return null;
    }
}



/**
 * 获取 Illustrator 中项目的树状结构
 * @param {Array} items - 项目数组
 * @return {string} 树状结构字符串
 */
function getItemTreeStructure(items) {
    var result = [];

    // 遍历所有项目
    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        // 获取项目的完整层级路径
        var path = getItemPath(item);

        // 将路径转换为树状结构格式
        var treeLine = convertPathToTree(path);
        result.push(treeLine);
    }

    return result.join('\n');
}

/**
 * 获取项目的完整层级路径
 * @param {Object} item - Illustrator 项目
 * @return {Array} 层级路径数组
 */
function getItemPath(item) {
    var path = [];
    var current = item;

    // 从当前项目向上遍历到文档根层级
    while (current && current.parent) {
        path.unshift(current); // 添加到数组开头
        current = current.parent;
    }

    return path;
}

/**
 * 将路径转换为树状结构格式
 * @param {Array} path - 层级路径数组
 * @return {string} 树状结构行
 */
function convertPathToTree(path) {
    var treeLine = '';

    for (var i = 0; i < path.length; i++) {
        var item = path[i];
        var isLast = i === path.length - 1;

        // 根据层级深度添加缩进
        if (i > 0) {
            treeLine += '⟼⟼';
        }

        // 添加项目标识和名称
        if (item.typename === 'Layer') {
            treeLine += '📃' + item.name;
        } else if (item.typename === 'GroupItem') {
            treeLine += '⟼' + item.name;
        } else {
            treeLine += '⟼⟼' + item.name;
        }

        if (!isLast) {
            treeLine += '\n';
        }
    }

    return treeLine;
}



/**
 * 获取树的最大深度
 */
function getMaxDepth(node) {
    var max = 0;
    for (var key in node) {
        if (node.hasOwnProperty(key)) {
            var currentDepth = 1 + getMaxDepth(node[key].children);
            if (currentDepth > max) max = currentDepth;
        }
    }
    return max;
}

 

/**
 * 递归构建树状结构（修复版本）
 */
function buildTree(node, path, depth,item,showChild) {
    if (depth >= path.length) return;

    var currentItem = path[depth];
    
    // 使用更唯一的键：包含父级信息 + 项目索引
    var parentInfo = depth > 0 ? path[depth-1].name + ':' : 'root:';
    var itemIndex = getItemIndexInParent(currentItem);
    var key = currentItem.typename + ':' + currentItem.name + ':' + itemDetail(currentItem, 'uniqueID') + ':' + parentInfo + itemIndex;

    if (!node[key]) {
        node[key] = {
            item: currentItem,
            children: {}
        };
    }

    if(showChild==false && currentItem==item){return }
    buildTree(node[key].children, path, depth + 1,item,showChild);
}

/**
 * 获取项目在父级中的索引
 */
function getItemIndexInParent(item) {
    if (!item.parent) return 0;
    
    var parent = item.parent;
    var pageItems = parent.pageItems || parent.layers || parent.groupItems;
    
    if (!pageItems) return 0;
    
    for (var i = 0; i < pageItems.length; i++) {
        if (pageItems[i] === item) {
            return i;
        }
    }
    return 0;
}



/**
 * 生成带锁定/可见状态的树状结构
 * 可选参数：
 *   style: "emoji" | "ascii" | "basic"
 *   indentChar: 缩进字符，默认 "＿"（全角下划线）
 *   showBothIcons: 是否显示锁+隐藏两个图标
 */
function getOptimizedItemTree(items, options) {
    options = options || {};
    var style = options.style || "emoji";
    var indentChar = options.indentChar || "＿";
    var showBothIcons = options.showBothIcons !== false; // 默认 true
    var showItemType = options.showItemType !== false;
    var showDocName= options.showDocName !== false;
    var showChild = options.showChild !== false; // 默认 true（显示子项）


    // 构建树状结构
    var rootStructure = {};
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var path = getItemPath(item);
        buildTree(rootStructure, path, 0,item,showChild);
    }

    // 计算最大深度
    var maxDepth = getMaxDepth(rootStructure);

    // 渲染树结构
    var result = [];
    generateTreeDisplay(rootStructure, result, 0, style, indentChar, showBothIcons, maxDepth,showItemType,showDocName, showChild);
    return result.join("\n");
}


/**
 * 递归生成树状显示
 */
function generateTreeDisplay(node, result, depth, style, indentChar, showBothIcons, maxDepth,showItemType,showDocName, showChild) {
    var keys = [];
    for (var key in node) if (node.hasOwnProperty(key)) keys.push(key);

    for (var i = 0; i < keys.length; i++) {
        var key = keys[i];
        var currentNode = node[key];
        var item = currentNode.item;

        var displayName = getDisplayName(item, depth, style, indentChar, showBothIcons, maxDepth,showItemType,showDocName);
         // 如果需要显示，则加入结果
        if (displayName !== false) {
            result.push(displayName);
        }

        // 子节点
        generateTreeDisplay(currentNode.children, result, depth + 1, style, indentChar, showBothIcons, maxDepth,showItemType,showDocName);
    }
}

/**
 * 单项显示（图标在前 + 自动对齐）
 */
function getDisplayName(item, depth, style, indentChar, showBothIcons, maxDepth, showItemType,showDocName) {
    if(typeof showDocName !='undefined' && showDocName === false && item.typename==="Document"){
        return false ;
    }
    var baseName = item.name || "未命名";
    // var lockIcon = item.locked === true ? "🔒" : "⚪️";
    // var visibleIcon = (item.typename=="Layer"?item.visible===false:item.hidden ===true  ) ? "🙈" : "⚪️";
    var lockIcon = item.locked === true ? "\uD83D\uDD12" : "\u26AA\uFE0F";
    var visibleIcon = (item.typename == "Layer" ? item.visible === false : item.hidden === true) ? "\uD83D\uDE48" : "\u26AA\uFE0F";
    var iconPart = showBothIcons ? (  visibleIcon+lockIcon) : lockIcon;

    // 统一缩进宽度：每层 2 个全角单位
    var indentPerLevel = 2;
    var indentWidth = indentPerLevel * depth;
    var indentPart = Array(indentWidth + 1).join(indentChar);

    // 对齐控制：补齐到最大深度宽度
    var maxWidth = indentPerLevel * maxDepth;
    var paddingWidth = maxWidth - indentWidth;
    var paddingPart = Array(paddingWidth + 1).join(indentChar);

    // 拼接显示
    return iconPart + indentPart + "- " + baseName + (showItemType?' : '+ item.typename:'') ; // + paddingPart;
}


function testing(){
    // var ids = allItemUniqueIDs();
    // uiDialog('confirm','ids 有： \n'+ids.join('\n'));

     var items = arrayLikeToArray(app.activeDocument.pageItems);
    
    // var items = arrayLikeToArray(app.activeDocument.selection);

    uiDialog('confirm','底层 items 有： \n'+items.filter(function(item){
            return (item.typename == 'GroupItem'&& item.pageItems.length==1)||(item.typename != 'GroupItem'&& item.typename != 'Layer');
        }).map(function(item){
            return item.name;
        }).join('\n'));

    var tree = getOptimizedItemTree(items,{
        style: "emoji",      // 可选 "emoji" | "ascii" | "basic"
        indentChar:"ㅤ ",// "＿",    // 使用全角下划线作为缩进
        showBothIcons: true,  // 显示锁+隐藏两个图标
        showItemType: true,
        showDocName:true ,
    });
    uiDialog('confirm','树状 ： \n'+tree); // 或者 $.writeln(tree);
    // alert(tree);
}

