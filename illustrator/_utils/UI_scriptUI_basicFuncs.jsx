

/** 
 * 根据按钮文本内容自动调整按钮宽度
 * @param {Object} btn - ScriptUI按钮对象
 */
function autoSizeButton(btn) {
    var text = btn.text;
    var gfx = btn.graphics;
    var size = gfx.measureString(text);
    btn.preferredSize.width = size[0] + 10; // 加点 padding
}


/**
 * ScriptUI 尺寸动画函数
 * @param {Window | Panel | Group | EditText} ctrl - 需要动画的控件
 * @param {Array} targetSizeArray - [width, height] 目标尺寸
 * @param {Number} step - 每步改变的像素数
 * @param {Number} stepDuration - 每步持续时间（毫秒）
 */
function animateSize(win, ctrl, targetSizeArray, step, stepDuration) {
    var currentWidth = ctrl.size[0];
    var currentHeight = ctrl.size[1];
    var targetWidth = targetSizeArray[0];
    var targetHeight = targetSizeArray[1];

    // 计算宽高变化方向
    var stepW = currentWidth < targetWidth ? Math.abs(step) : -Math.abs(step);
    var stepH = currentHeight < targetHeight ? Math.abs(step) : -Math.abs(step);

    while (true) {
        var doneW = false, doneH = false;

        // 更新宽度
        if ((stepW > 0 && currentWidth < targetWidth) || (stepW < 0 && currentWidth > targetWidth)) {
            currentWidth += stepW;
            if ((stepW > 0 && currentWidth > targetWidth) || (stepW < 0 && currentWidth < targetWidth)) currentWidth = targetWidth;
        } else {
            doneW = true;
        }

        // 更新高度
        if ((stepH > 0 && currentHeight < targetHeight) || (stepH < 0 && currentHeight > targetHeight)) {
            currentHeight += stepH;
            if ((stepH > 0 && currentHeight > targetHeight) || (stepH < 0 && currentHeight < targetHeight)) currentHeight = targetHeight;
        } else {
            doneH = true;
        }

        // 应用尺寸并刷新布局
        ctrl.size = [currentWidth, currentHeight];
        ctrl.parent.layout.layout(true);

        if (doneW && doneH) break; // 完成

        $.sleep(stepDuration);
        win.update(); // 刷新窗口
        win.layout.layout(true); // 重新布局
    }
}


function edittextPlaceholder(edittext, placeholderText) {
    edittext.text = placeholderText;
    var firstFocus = true; // 标记第一次点击

    edittext.onActivate = function () {
        if (firstFocus) {
            edittext.text = "";
            firstFocus = false;
        }
    }
    edittext.onDeactivate = function () {
        if (edittext.text === "") {
            edittext.text = placeholderText;
            firstFocus = true;
        }
    }
}


function deepEachPageItemSetProp(container, prop, value) {
    // 参数验证
    if (!container || !container.pageItems) {
        return;
    }

    // 使用迭代而不是递归来避免栈溢出
    var stack = [];

    // 初始化栈
    for (var i = 0; i < container.pageItems.length; i++) {
        stack.push(container.pageItems[i]);
    }

    // 迭代处理
    while (stack.length > 0) {
        var item = stack.pop();

        try {
            // 设置属性
            if (item && item.hasOwnProperty(prop)) {
                item[prop] = value;
            }

            // 如果有子项，添加到栈中
            if (item.pageItems && item.pageItems.length > 0) {
                for (var j = 0; j < item.pageItems.length; j++) {
                    stack.push(item.pageItems[j]);
                }
            }
        } catch (e) {
            // 记录错误但继续执行
            log("Error processing item: " + e.toString());
        }
    }
}