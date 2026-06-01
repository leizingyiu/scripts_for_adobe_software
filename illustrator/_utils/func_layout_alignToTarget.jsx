// 已测试 ✅

 
/**
 * Align children of a group to a target item based on specified alignment
 * @param {GroupItem} group - The group containing children to align
 * @param {PageItem} targetItem - The target item to align to
 * @param {String} alignment - Alignment direction: "left", "right", "top", or "bottom"
 */


function alignChildrenToTarget(group, targetItem, alignment) {
    if (!group || !group.hasOwnProperty('pageItems') || group.pageItems.length <= 1) {
        return;
    }

    if (!targetItem) {
        return;
    }

    var items = group.pageItems;
    var bounds = [];

    // 获取所有可见子项的边界框
    for (var i = 0; i < items.length; i++) {
        if (items[i].hidden) {
            continue;
        }
        bounds.push(items[i].geometricBounds);
    }

    if (bounds.length <= 1) {
        return;
    }

    // 获取目标对象的边界框
    var targetBounds = targetItem.geometricBounds;

    // Helper: 对齐所有对象到目标对象的指定边
    function alignItemsToTarget(targetValue, index) {
        for (var i = 0; i < items.length; i++) {
            if (items[i].hidden) {
                continue;
            }
            var delta = targetValue - bounds[i][index];
            if (index === 0 || index === 2) { // X轴：left/right
                items[i].translate(delta, 0);
            } else { // Y轴：top/bottom
                items[i].translate(0, delta);
            }
        }
    }

    switch (alignment.toLowerCase()) {
        case "left":
            // 对齐所有对象的左边缘到目标对象的左边缘
            alignItemsToTarget(targetBounds[0], 0);
            break;

        case "right":
            // 对齐所有对象的右边缘到目标对象的右边缘
            alignItemsToTarget(targetBounds[2], 2);
            break;

        case "top":
            // 对齐所有对象的上边缘到目标对象的上边缘
            alignItemsToTarget(targetBounds[1], 1);
            break;

        case "bottom":
            // 对齐所有对象的下边缘到目标对象的下边缘
            alignItemsToTarget(targetBounds[3], 3);
            break;

        default:
            alert("无效对齐选项。请使用 'left', 'right', 'top', 或 'bottom'。");
            break;
    }
}


/**
 * Testing function - aligns selected group's children to the second child as target item, using left alignment
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
    
    var selection = doc.selection[0];
    
    if (selection.typename !== "GroupItem") {
        alert("请选择一个组对象。");
        return;
    }
    
    if (selection.pageItems.length < 2) {
        alert("请选择至少包含2个元素的组。");
        return;
    }
    
    var targetItem = selection.pageItems[1]; // Second child as target
    
    alignChildrenToTarget(selection, targetItem, "left");
    
    alert("已将组内元素向第二个元素左对齐完成。");
}

/**
 * Testing2 function - aligns selected group's children to the second child as target item, using left alignment
 */
function testing2() {
    if (app.documents.length === 0) {
        alert("请打开一个文档。");
        return;
    }
    
    var doc = app.activeDocument;
    
    if (doc.selection.length === 0) {
        alert("请选择一个组。");
        return;
    }
    
    var selection = doc.selection;
    
    // if (selection.typename !== "GroupItem") {
    //     alert("请选择一个组对象。");
    //     return;
    // }
    
    // if (selection.pageItems.length < 2) {
    //     alert("请选择至少包含2个元素的组。");
    //     return;
    // }
    

    for(var i=0;i<selection.length;i++){

        var group = selection[i];
        var targetItem = group.pageItems[1]; // Second child as target
        alignChildrenToTarget(group, targetItem, group.name );
    }
    
    

}

// Run the testing function
// testing2();