
var targetTextFrameName = '1:someText {_(框) `{_ y} y {_ m} m {_ d} d 范本文本范本文本范本文本范本文本范本文本范本文本范本文本范本文本范本文本范本' +
        '文本`} text'; //"YourTextFrameName"; // 替换为目标文本框名称
var testContent = "exampleText"; 
var testContent2 = 'example\ntext';

// ===========================

function fillText_testing() {
    if (app.documents.length === 0) {
        alert("请打开一个文档。");
        return;
    }

    var doc = app.activeDocument;

    function findTextFrameByName(doc, name) {
        var frames = doc.textFrames;
        for (var i = 0; i < frames.length; i++) {
            if (frames[i].name == name) {
                return frames[i];
            }
        }
        return null;
    }
    var textFrame = findTextFrameByName(doc, targetTextFrameName);

    if (!textFrame) {
        alert("未找到名称为 " + targetTextFrameName + " 的文本框。");
        return;
    }

    fillTxtInFrame(textFrame, testContent);

    // var input = prompt("请输入要填入的文本内容：", ''); testContent2 =
    // input.split('\n')[0].split('\t')[1].replace('\\n','\n') ;
    // fillAndKeepLinesForTxtAvatar( textFrame,testContent2);

}

/** 填充文本框 */
function fillText(item, content) {
    try {
        item.story.textRange.contents = content;
    } catch (e) {
        throw new Error('出现错误：\n '+e+'\n fillText()');
    }
}

/** 通过用内容对象中的值替换占位符来填充模板中的文本内容
 * @param {Object} item - 包含模板的文本项对象,必须具有name和contents属性
 * @param {Object} contentObj - 包含用于模板替换的键值对的对象
 *
 * @description
 * 函数的工作原理:
 * 1. 从item名称中提取反引号(`)之间的模板
 * 2. 在模板中查找所有格式为{key}的占位符
 * 3. 用contentObj中的对应值替换每个占位符
 * 4. 将新内容设置为item.contents
 *
 * @example
 * // item.name = "text`Hello {name}!`"
 * // contentObj = {name: "John"}
 * // 函数调用后: item.contents = "Hello John!"
 */

function fillTextInTemplet(item, contentObj) {
    var templet = item
        .name
        .replace(/.*?`([^`]*)`.*/g, '$1');
    var newContent = templet;
    var ks = templet.match(/{([^}]*)}/g);
    if (ks) {
        for (var i = 0; i < ks.length; i++) {
            var k = ks[i].replace(/[{}_]/g, '');
            if (contentObj[k]) {
                newContent = newContent.replace(ks[i], contentObj[k]);
            }
        }
    }
    fillText(item, newContent);
}


/**  fillTxtInFrame
 *
 * @param {TextFrame} textFrame
 * @param {String} content
 * @returns
 *
 * @description
 * 1. 填充文本内容
 * 2. 自动调整字体大小，使文本框内不溢出
 * 3. 返回最终字体大小
 * 4. 如果溢出，返回溢出字体大小和中间值
 * 5. 如果未溢出，返回最终字体大小
 */
function fillTxtInFrame(textFrame, content) {

    var upFontSize2FillTxtFrame = function (tf) {
        if (!(tf instanceof TextFrame)) {
            throw new Error("传入的不是 TextFrame");
        }

        var chAttr = tf.textRange.characterAttributes;
        var originalLineHeightEm = tf.textRange.characterAttributes.leading / tf.textRange.characterAttributes.size;

        // 初始字号
        var size = 1;
        var maxSize = 500;
        var step = 0.3;
        var maxTry = 100;
        var _try = 0;
        chAttr.size = size;
        chAttr.leading = chAttr.size * originalLineHeightEm;

        var prevRowCount = tf.lines.length;
        var prevSize = size;

        var lastSafeSize = size; // 记录最后一个安全的字号（行数增加前的那个）
        var maxRowCount = tf.lines.length; // 记录最大行数
        var overflowDetected = false;

        while (size < maxSize) {
            _try++;
            size += step;
            chAttr.size = size;
            chAttr.leading = chAttr.size * originalLineHeightEm;

            var currentRowCount = tf.lines.length;

            // 更新最大行数
            if (currentRowCount > maxRowCount) {
                maxRowCount = currentRowCount;
                _try = 0; // 重置计数器
                lastSafeSize = prevSize; // 记录行数增加前的安全点
            }

            // 检测到文本溢出（行数减少了）
            if (currentRowCount < maxRowCount) {
                overflowDetected = true;

                // 使用 lastSafeSize 和 当前 size 的中间值作为最终安全字号
                var finalSize = (lastSafeSize + size) / 2;

                // 保守一点，可以再减 0.1 避免临界误差
                finalSize -= 0.1;
                if (finalSize < 1) 
                    finalSize = 1;
                
                chAttr.size = finalSize;
                chAttr.leading = chAttr.size * originalLineHeightEm;

                // testingAlert([     'Overflow Detected - Using Binary Midpoint',
                // 'lastSafeSize: ' + lastSafeSize,     'overflowSize: ' + size,     'midpoint:
                // ' + ((lastSafeSize + size) / 2),     'finalSize (adjusted): ' + finalSize,
                //  'maxRowCount: ' + maxRowCount,     'currentRowCount: ' + currentRowCount
                // ].join('\n'));

                return {
                    maxRowCount: maxRowCount,
                    overflowSize: size,
                    finalSize: finalSize,
                    lastSafeSize: lastSafeSize,
                    midpoint: (lastSafeSize + size) / 2
                };
            }

            // 行数稳定且超过最大尝试次数
            if (currentRowCount === prevRowCount && _try > maxTry) {
                // 已经达到最佳大小 testingAlert(['Stable size reached',     'size: ' + size,
                // 'maxRowCount: ' + maxRowCount,     'currentRowCount: ' + currentRowCount,
                // 'finalSize: ' + size ].join('\n'));

                return {maxRowCount: maxRowCount, finalSize: size};
            }

            prevRowCount = currentRowCount;
            prevSize = size;
        }

        // 如果循环结束仍未溢出，返回最后尝试的大小
        return {maxRowCount: maxRowCount, finalSize: size};
    }
    fillText(textFrame, content);

    upFontSize2FillTxtFrame(textFrame);
}

/** 适用于文本头像的文本框，要求行数不变,自动缩小字体,使用段后间距作为行距
    * @param {TextFrame} textFrame - 目标文本框对象
    * @param {String} content - 要填充的文本内容
    *
    * 功能说明:
    * - 将指定内容填充到文本框中
    * - 保持原始行高比例不变
    * - 如果文本行数不匹配,会逐步缩小字体直到行数匹配
    * - 最小字体大小限制为4pt
    * - 每次调整字号减少0.1pt
    * 内部辅助函数:
    * - txtRngSmallerStepFn: 按指定步长减小字号并保持行高比例
    * - isSameLines: 检测文本行数是否与内容行数匹配
*/
function fillAndKeepLinesForTxtAvatar(textFrame, content) {
    // alert(['fillAndKeepLinesForTxtAvatar',textFrame.name,content].join('\n'));

    var textRange = textFrame.story.textRange;
    var originalLineHeightEm = textFrame.paragraphs[0].paragraphAttributes.spaceAfter / textRange.characterAttributes.size;
    textRange.contents = content;
    function isSameLines(textRange, content) {
        return textRange.lines.length == content
            .split('\n')
            .length;
    }
    var txtRngSmallerStepFn = function (textRange, smallStepNum) {
        textRange.characterAttributes.size -= smallStepNum;
        for (var pi = 0; pi < textRange.parent.paragraphs.length; pi++) {
            textRange.parent.paragraphs[pi].paragraphAttributes.spaceAfter = textRange.characterAttributes.size * originalLineHeightEm;
        }
    }
    if (!isSameLines(textRange, content)) {
        while (!isSameLines(textRange, content) && textRange.characterAttributes.size > 4) {
            txtRngSmallerStepFn(textRange, 0.1);
        }
    }
}


/**
 * 在原 TextFrame 上安全替换文本，自动调整宽度且保留样式
 * @param {TextFrame} item - 要操作的文本对象
 * @param {String} newString - 要填入的文本
 */
function fillAreaTxtWider(item, newString) {
    /**
     * 判断文本样式是否一致
     */
    var isUniformStyle = function (tr) {
        var ca = tr.characterAttributes;
        var pa = tr.paragraphAttributes;

        for (var i = 0; i < tr.characters.length; i++) {
            if (!compareCharAttr(tr.characters[i].characterAttributes, ca)) return false;
        }
        for (var j = 0; j < tr.paragraphs.length; j++) {
            if (!compareParaAttr(tr.paragraphs[j].paragraphAttributes, pa)) return false;
        }
        return true;
    }

    /**
     * 克隆字符样式
     */
    var cloneCharAttrs = function (src) {
        var obj = {};
        obj.size = src.size;
        obj.textFont = src.textFont;
        obj.fillColor = src.fillColor;
        obj.tracking = src.tracking;
        obj.horizontalScale = src.horizontalScale;
        obj.verticalScale = src.verticalScale;
        obj.kerning = src.kerning;
        obj.leading = src.leading;
        obj.opacity = src.opacity;
        return obj;
    }

    /**
     * 克隆段落样式
     */
    var cloneParaAttrs = function (src) {
        var obj = {};
        obj.justification = src.justification;
        obj.leftIndent = src.leftIndent;
        obj.rightIndent = src.rightIndent;
        obj.spaceBefore = src.spaceBefore;
        obj.spaceAfter = src.spaceAfter;
        obj.firstLineIndent = src.firstLineIndent;
        return obj;
    }

    /**
     * 应用字符样式
     */
    var applyCharAttrs = function (target, src) {
        try {
            target.size = src.size;
            target.textFont = src.textFont;
            target.fillColor = src.fillColor;
            target.tracking = src.tracking;
            target.horizontalScale = src.horizontalScale;
            target.verticalScale = src.verticalScale;
            target.kerning = src.kerning;
            target.leading = src.leading;
            target.opacity = src.opacity;
        } catch (e) { }
    }

    /**
     * 应用段落样式
     */
    var applyParaAttrs = function (target, src) {
        try {
            target.justification = src.justification;
            target.leftIndent = src.leftIndent;
            target.rightIndent = src.rightIndent;
            target.spaceBefore = src.spaceBefore;
            target.spaceAfter = src.spaceAfter;
            target.firstLineIndent = src.firstLineIndent;
        } catch (e) { }
    }

    /**
     * 比较字符样式
     */
    var compareCharAttr = function (a, b) {
        if (!a || !b) return false;
        try {
            if (a.size !== b.size) return false;
            if (a.textFont !== b.textFont) return false;
            if (!compareColor(a.fillColor, b.fillColor)) return false;
            if (a.tracking !== b.tracking) return false;
            if (a.horizontalScale !== b.horizontalScale) return false;
            if (a.verticalScale !== b.verticalScale) return false;
            if (a.opacity !== b.opacity) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 比较段落样式
     */
    var compareParaAttr = function (a, b) {
        if (!a || !b) return false;
        try {
            if (a.justification !== b.justification) return false;
            if (a.leftIndent !== b.leftIndent) return false;
            if (a.rightIndent !== b.rightIndent) return false;
            if (a.spaceBefore !== b.spaceBefore) return false;
            if (a.spaceAfter !== b.spaceAfter) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 比较颜色
     */
    var compareColor = function (c1, c2) {
        if (!c1 || !c2) return false;
        if (c1.typename !== c2.typename) return false;
        try {
            if (c1.red !== c2.red) return false;
            if (c1.green !== c2.green) return false;
            if (c1.blue !== c2.blue) return false;
            return true;
        } catch (e) {
            return false;
        }
    }
    if (!item || item.typename !== "TextFrame") {
        throw new Error ("fillAreaTextWider() : 错误：传入的对象不是 TextFrame");
        return;
    }

    // 如果是点文字
    if (item.kind !== TextType.AREATEXT) {
        item.contents = newString;
        return item ;
    }

    var tr = item.textRange;
    if (!isUniformStyle(tr)) {
        throw new Error ("fillAreaTextWider() : 错误：文本框中存在不同样式或颜色，无法安全替换。");
        return item ;
    }

    // 记录当前样式
    var charAttrs = cloneCharAttrs(tr.characterAttributes);
    var paraAttrs = cloneParaAttrs(tr.paragraphAttributes);

    // 临时点文本用于测量新文字宽度
    var doc = app.activeDocument;
    var tmp = doc.textFrames.add();
    tmp.contents = newString;
    applyCharAttrs(tmp.textRange.characterAttributes, charAttrs);
    tmp.position = item.position;

    var tmpWidth = tmp.width;
    var boxWidth = item.width;

    // === ⚙️ 关键步骤：调整宽度前先清空文本，避免样式断裂 ===
    item.contents = ""; // 避免 AI 重排导致样式丢失
    var oldWidth = item.width;

    if (tmpWidth > boxWidth) {
        item.width = tmpWidth;
    } else {
        item.width = boxWidth; // 保持原宽度
    }

    // 重新设置样式 + 内容
    applyCharAttrs(item.textRange.characterAttributes, charAttrs);
    applyParaAttrs(item.textRange.paragraphAttributes, paraAttrs);
    item.contents = newString;

    // 清理临时对象
    tmp.remove();

    return item;
}

// testing();