// 按排列方向前两个的位置，排列剩余元素，如：left：左item左 到 右item左 ；
// 已测试 ✅
// testing();


//TODO ： 添加上右下左的平均分布 如： left： 最左最右固定，剩余左到左、左到左 平均；


#target = "illustrator"

/**  * Distribute children of a group in specified direction with calculated spacing
 * Spacing is calculated from the distance between first two items, and applied to the rest
 * @param {GroupItem} group - The group containing children to distribute
 * @param {String} direction - Distribution direction: "left", "right", "top", or "bottom"
 */
function distributeChildren(group, direction) {

    if (!group || !group.hasOwnProperty('pageItems') || group.pageItems.length < 2) {
        throw "Invalid group. Please select a group with at least 2 children.";
        return;
    }

    var items = group.pageItems;
    var _items = [];      // 只包含 visible 的 item
    var bounds = [];      // 对应 _items 的 geometricBounds

    // 收集所有非隐藏的子项及其边界
    for (var i = 0; i < items.length; i++) {
        if (items[i].hidden) {
            continue;
        }
        _items.push(items[i]);
        bounds.push(items[i].geometricBounds);
    }

    // 如果可见对象少于2个，无法分布
    if (_items.length < 2) {
        throw "Invalid group. Please select a group with at least 2 children.";
        return;
    }

    // sortedIndices: [0, 1, 2, ..., _items.length - 1]
    var sortedIndices = [];
    for (var i = 0; i < _items.length; i++) {
        sortedIndices.push(i);
    }

    // 根据方向排序
    var dir = direction.toLowerCase();
    switch (dir) {
        case "left":
            sortedIndices.sort(function(a, b) {
                return bounds[a][0] - bounds[b][0]; // left edge
            });
            break;
        case "right":
            sortedIndices.sort(function(a, b) {
                return bounds[b][2] - bounds[a][2]; // right edge (larger x = more right)
            });
            break;
        case "top":
            sortedIndices.sort(function(a, b) {
                return bounds[b][1] - bounds[a][1]; // top edge (larger y = higher)
            });
            break;
        case "bottom":
            sortedIndices.sort(function(a, b) {
                return bounds[a][3] - bounds[b][3]; // bottom edge (smaller y = lower)
            });
            break;
        default:
            alert("Invalid direction. Use 'left', 'right', 'top', or 'bottom'.");
            return;
    }

    // 确定边缘索引和方向
    var edgeIndex, isVertical;
    switch (dir) {
        case "left":
            edgeIndex = 0; isVertical = false;
            break;
        case "right":
            edgeIndex = 2; isVertical = false;
            break;
        case "top":
            edgeIndex = 1; isVertical = true;
            break;
        case "bottom":
            edgeIndex = 3; isVertical = true;
            break;
    }

    // 计算前两个可见对象的间距
    var spacing = bounds[sortedIndices[1]][edgeIndex] - bounds[sortedIndices[0]][edgeIndex];

 
    // 从第3个开始分布
    for (var i = 2; i < sortedIndices.length; i++) {
        var prevEdge = bounds[sortedIndices[i - 1]][edgeIndex];
        var currentEdge = bounds[sortedIndices[i]][edgeIndex];
        var targetPos = prevEdge + spacing;
        var delta = targetPos - currentEdge;

        // 移动对象
        if (isVertical) {
            _items[sortedIndices[i]].translate(0, delta);
        } else {
            _items[sortedIndices[i]].translate(delta, 0);
        }

        // 更新 bounds（用于后续计算）
        var width = bounds[sortedIndices[i]][2] - bounds[sortedIndices[i]][0];
        var height = bounds[sortedIndices[i]][3] - bounds[sortedIndices[i]][1];

        // 更新移动后的边界
        if (edgeIndex === 0) { // left
            bounds[sortedIndices[i]][0] = targetPos;
            bounds[sortedIndices[i]][2] = targetPos + width;
        } else if (edgeIndex === 1) { // top
            bounds[sortedIndices[i]][1] = targetPos;
            bounds[sortedIndices[i]][3] = targetPos + height;
        } else if (edgeIndex === 2) { // right
            bounds[sortedIndices[i]][2] = targetPos;
            bounds[sortedIndices[i]][0] = targetPos - width;
        } else if (edgeIndex === 3) { // bottom
            bounds[sortedIndices[i]][3] = targetPos;
            bounds[sortedIndices[i]][1] = targetPos - height;
        }
    }
 }


/**
 * Testing function - distributes selected group's children to the left
 */
function testing() {
    if (app.documents.length === 0) {
        alert("请打开一个文档。");
        return;
    }
    
    var doc = app.activeDocument;
    
    if (doc.selection.length === 0) {
        alert("请选择一个组。");
        return;
    }
    
    var directions = ['left','right','top','bottom'];

    

    for(var i = 0 ;i<doc.selection.length;i++){

        var group = doc.selection[i];
        
        if (group.typename !== "GroupItem") {
            alert("请选择一个组对象。");
            return;
        }
        
        if (group.pageItems.length < 3) {
            alert("请至少选择包含3个元素的组。");
            return;
        }
        // alert("正在处理组：" + group.name+ "..."+directions[i]);
        
        distributeChildren(group, directions[i]  );
    }
    
    alert("已将组内元素按左方向排列完成。");
}

// Run the testing function
