// 按排列方向前两个之间的距离，排列剩余元素，如：left：左item右 到 右item左 ；
// 已测试 ✅
// testing();

// TODO ： 添加两侧平均分布
// TODO ： 添加指定对象固定，使用原来剩余距离平均分布

/**
 * ES3 JSX script for Adobe Illustrator
 * Space children items within a group based on direction and spacing
 */

/**
 * Space children of a group in specified direction with calculated spacing
 * Spacing is calculated from the distance between first two items (including their widths),
 * and applied to the rest
 * @param {GroupItem} group - The group containing children to space
 * @param {String} direction - Spacing direction: "left", "right", "top", or "bottom"
 */
function spaceChildren(group, direction) {

    try{ 

        if (!group || !group.hasOwnProperty('pageItems') || group.pageItems.length <= 2) {
            return;
        }
        
        var items = group.pageItems;
        var bounds = [];
        


        // Get bounds of all children
        for (var i = 0; i < items.length; i++) {
            if(items[i].hidden) {
                continue;
            }
            bounds.push(items[i].geometricBounds);
        }
        

        // Sort items based on direction
        var sortedIndices = [];
        for (var i = 0; i < items.length; i++) {
            if(items[i].hidden) {
                continue;
            }
            sortedIndices.push(i);
        }





        switch (direction.toLowerCase()) {
                case "left":
                // Sort by x position (left edge)
                sortedIndices.sort(function(a, b) {
                    return bounds[a][0] - bounds[b][0];
                });
                break;
                
                case "right":
                // Sort by x position (right edge)
                sortedIndices.sort(function(a, b) {
                    return bounds[b][2] - bounds[a][2];
                });
                break;

                case "up":
                case "top":
                // Sort by y position (top edge)
                sortedIndices.sort(function(a, b) {
                    return bounds[b][1] - bounds[a][1]; // Top position, larger y = higher
                });
                break;
                
                case "down":
                case "bottom":
                // Sort by y position (bottom edge)
                sortedIndices.sort(function(a, b) {
                    return bounds[a][3] - bounds[b][3]; // Bottom position, smaller y = lower
                });
                break;
            default:
                alert("Invalid direction option. Please use 'left', 'right', 'top', or 'bottom'.");
                return;
        }
        

        // Calculate spacing based on first two items
        var spacing = 0;
        var edgeIndex1, edgeIndex2, sizeIndex, isVertical;
        


        switch (direction.toLowerCase()) {
            case "left":
                // Spacing = distance from first item's right edge to second item's left edge
                edgeIndex1 = 2; // First item's right edge
                edgeIndex2 = 0; // Second item's left edge
                sizeIndex = 0;  // X-axis
                isVertical = false;
                break;
            case "right":
                // Spacing = distance from first item's left edge to second item's right edge
                edgeIndex1 = 0; // First item's left edge
                edgeIndex2 = 2; // Second item's right edge
                sizeIndex = 0;  // X-axis
                isVertical = false;
                break;

            case "up":
            case "top":
                // Spacing = distance from first item's bottom edge to second item's top edge
                edgeIndex1 = 3; // First item's bottom edge
                edgeIndex2 = 1; // Second item's top edge
                sizeIndex = 1;  // Y-axis
                isVertical = true;
                break;
            case "down": 
            case "bottom":
                // Spacing = distance from first item's top edge to second item's bottom edge
                edgeIndex1 = 1; // First item's top edge
                edgeIndex2 = 3; // Second item's bottom edge
                sizeIndex = 1;  // Y-axis
                isVertical = true;
                break;
        }
        

        // Calculate spacing between first two items
        var firstEdge = bounds[sortedIndices[0]][edgeIndex1];
        var secondEdge = bounds[sortedIndices[1]][edgeIndex2];
        spacing = secondEdge - firstEdge;
        

        // Position each subsequent item based on previous item's position and size
        for (var i = 2; i < sortedIndices.length; i++) {
            var prevIndex = sortedIndices[i-1];
            var currentIndex = sortedIndices[i];
            
            // Calculate previous item's edge position
            var prevEdge;
            if (direction.toLowerCase() === "left") {
                // Use previous item's right edge
                prevEdge = bounds[prevIndex][2];
            } else if (direction.toLowerCase() === "right") {
                // Use previous item's left edge
                prevEdge = bounds[prevIndex][0];
            } else if (direction.toLowerCase() === "top") {
                // Use previous item's bottom edge
                prevEdge = bounds[prevIndex][3];
            } else if (direction.toLowerCase() === "bottom") {
                // Use previous item's top edge
                prevEdge = bounds[prevIndex][1];
            }
            
            // Calculate current item's edge position
            var currentEdge = bounds[currentIndex][edgeIndex2];
            
            // Calculate target position
            var targetPos = prevEdge + spacing;
            var delta = targetPos - currentEdge;
            
            // Move the item
            if (isVertical) {
                items[currentIndex].translate(0, delta);
            } else {
                items[currentIndex].translate(delta, 0);
            }
            
            // Update bounds for next calculation
            var width = bounds[currentIndex][2] - bounds[currentIndex][0];
            var height = bounds[currentIndex][3] - bounds[currentIndex][1];
            
            bounds[currentIndex][edgeIndex2] = targetPos;
            if (edgeIndex2 === 0) { // Left edge updated
                bounds[currentIndex][2] = targetPos + width;
            } else if (edgeIndex2 === 1) { // Top edge updated
                bounds[currentIndex][3] = targetPos + height;
            } else if (edgeIndex2 === 2) { // Right edge updated
                bounds[currentIndex][0] = targetPos - width;
            } else if (edgeIndex2 === 3) { // Bottom edge updated
                bounds[currentIndex][1] = targetPos - height;
            }
        }


    }catch(e){
        // alert("Error: " + e +"\n" + e.stack);
        throw new Error(e + "\n" + e.stack);
        }

}

/**
 * Testing function - spaces selected group's children to the left
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
    
    // var selection = doc.selection[0];
    
    // if (selection.typename !== "GroupItem") {
    //     alert("请选择一个组对象。");
    //     return;
    // }
    
    // if (selection.pageItems.length < 3) {
    //     alert("请至少选择包含3个元素的组。");
    //     return;
    // }

    var selection = doc.selection;

    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename !== "GroupItem") {
            alert("请选择一个组对象。");
            return;
        }
        
        if (item.pageItems.length < 3) {
            alert("请至少选择包含3个元素的组。");
            return;
        }
        spaceChildren(item, item.name );
    }
    
    
    alert("已将组内元素按左方向设置间距完成。");
}

// Run the testing function
