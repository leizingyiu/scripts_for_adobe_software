// 已测试 ✅
 
/**
 * Align children of a group to specified direction
 * @param {GroupItem} group - The group containing children to align
 * @param {String} alignment - Alignment direction: "left", "right", "top", or "bottom"
 */

function alignChildren(group, alignment) {
    if (!group || !group.hasOwnProperty('pageItems') || group.pageItems.length <= 1) {
        return;
    }

    var items = group.pageItems;
    var bounds = [];
    var visibleItems = []; // 存储实际参与对齐的可见 item

    // Helper: 获取一个项目的对齐边界（考虑剪切蒙版）
    function getBoundsForAlignment(item) {
        // 检查是否是剪切蒙版组
        if (item.typename === 'GroupItem' && item.clipped && item.pageItems.length > 0) {
            var clipPaths = [];
            // 遍历组内寻找剪切路径
            for (var i = 0; i < item.pageItems.length; i++) {
                var subItem = item.pageItems[i];
                // 在 Illustrator 脚本中，剪切路径的 'clipping' 属性为 true
                // 并且它通常是组内的第一个对象，但为了鲁棒性，我们遍历查找
                if (subItem.clipping) {
                    clipPaths.push(subItem.geometricBounds);
                }
            }

            // 如果找到了剪切路径
            if (clipPaths.length > 0) {
                // 如果找到多个（虽然不常见但可以处理），需要合并边界
                // 这里的逻辑是获取所有剪切路径的最小包含矩形
                var minX = clipPaths[0][0];
                var maxY = clipPaths[0][1];
                var maxX = clipPaths[0][2];
                var minY = clipPaths[0][3];

                for (var i = 1; i < clipPaths.length; i++) {
                    minX = Math.min(minX, clipPaths[i][0]);
                    maxY = Math.max(maxY, clipPaths[i][1]);
                    maxX = Math.max(maxX, clipPaths[i][2]);
                    minY = Math.min(minY, clipPaths[i][3]);
                }
                
                // 返回合并后的边界 [minX, maxY, maxX, minY]
                return [minX, maxY, maxX, minY];
            }
        }
        
        // 普通对象或未找到剪切路径的组，使用对象的 geometricBounds
        return item.geometricBounds;
    }

    // --- 修改开始：获取所有可见子项的边界框 ---
    for (var i = 0; i < items.length; i++) {
        if (items[i].hidden) {
            continue;
        }
        // 使用新的辅助函数获取边界
        bounds.push(getBoundsForAlignment(items[i]));
        visibleItems.push(items[i]); // 记录参与对齐的对象
    }
    // --- 修改结束 ---

    if (bounds.length <= 1) {
        return;
    }
    
    // Helper: 获取指定索引的极值（min 或 max）
    function getExtremeValue(index, isMin) {
        var extreme = bounds[0][index];
        for (var i = 1; i < bounds.length; i++) {
            // 注意：几何边界的坐标是 [minX, maxY, maxX, minY]
            // index 0: minX, index 1: maxY (top), index 2: maxX, index 3: minY (bottom)
            if (isMin ? bounds[i][index] < extreme : bounds[i][index] > extreme) {
                extreme = bounds[i][index];
            }
        }
        return extreme;
    }

    // Helper: 对齐所有对象到指定边
    function alignItems(extremeValue, index) {
        for (var i = 0; i < visibleItems.length; i++) {
            // visibleItems[i] 对应 bounds[i]
            var currentBounds = bounds[i];
            
            // 计算平移量
            var delta = extremeValue - currentBounds[index];
            
            // X轴：left(0) / right(2)
            if (index === 0 || index === 2) { 
                visibleItems[i].translate(delta, 0);
            } 
            // Y轴：top(1) / bottom(3)
            else { 
                visibleItems[i].translate(0, delta);
            }
        }
    }

    switch (alignment.toLowerCase()) {
        case "left":
            var leftMost = getExtremeValue(0, true); // 最左边缘 (minX)
            alignItems(leftMost, 0);
            break;

        case "right":
            var rightMost = getExtremeValue(2, false); // 最右边缘 (maxX)
            alignItems(rightMost, 2);
            break;

        case "top":
            var topMost = getExtremeValue(1, false); // 最上边缘 (maxY)
            alignItems(topMost, 1);
            break;

        case "bottom":
            var bottomMost = getExtremeValue(3, true); // 最下边缘 (minY)
            alignItems(bottomMost, 3);
            break;

        default:
            alert("无效对齐选项。请使用 'left', 'right', 'top', 或 'bottom'。");
            break;
    }
}

/**
 * Testing function - aligns selected group's children to the left
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
    
    var selection = doc.selection;
    
    var directions=['left','right','top','bottom'];
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename !== "GroupItem") {
            alert("请选择一个组对象。");
            return;
        }

        alignChildren(selection[i], selection[i].name);

    }
}

 // testing();