#target illustrator
// #include "UI_scriptUI_basicFuncs.jsx" // animateSize
// #include "utils_JSON_AdobeSafe.jsx"


// 已测试 ✅
// testing();
// function testing() {
//     showExcelParserDialog({
//         "alert": function (msg) {
//             alert(JSON.stringify(msg)); return false;
//         },
//         "alert2": function (msg) {
//             return JSON.stringify(msg);
//         }
//     });
// }



function showExcelParserDialog(fnsDict) {
    /**
     * @param {Object} fnsDict - 函数字典
     * {btnName : function(thisData.currentResult){
     * 
     * }}
     */
    // this.data = {};
    var thisData = {
        currentResult: null,
        rawInput: "",
        parseMode: null,
        closeMsg: "确定要关闭吗？"
    };

    var win = new Window("dialog", "Excel to JSON Parser");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.preferredSize.width = 560;
    win.preferredSize.height = 600;

    // === 输入区 ===

    var firstFocus = true; // 标记第一次点击

    win.add("statictext", undefined, "请粘贴来自 Excel 的表格内容（Tab 分隔）:");
    var placeholderText = '请在这里粘贴内容，点击按钮进行解析';
    var inputBox = win.add("edittext", undefined, '', { multiline: true, scrolling: true });
    inputBox.preferredSize.height = 150;
    inputBox.margins = [10, 10, 10, 10];

    edittextPlaceholder(inputBox, placeholderText);



    inputBox.alignment = ["fill", "top"];

    // === 最大长度输入框 ===

    // === 操作区 ===
    var ctrlGroup = win.add("group");
    // ctrlGroup.alignment = "center";
    ctrlGroup.orientation = "row";
    ctrlGroup.margins = 0;
    ctrlGroup.spacing = 10;
    ctrlGroup.alignment = ["fill", "top"];

    var lenGroup = ctrlGroup.add("group");
    lenGroup.add("statictext", undefined, "文本最大显示长度：");
    lenGroup.spacing = 0;
    var lenInput = lenGroup.add("edittext", undefined, "15");
    lenInput.characters = 3;

    var btnAuto = ctrlGroup.add("button", undefined, "自动解析");
    var btnRow = ctrlGroup.add("button", undefined, "按行解析");
    btnRow.alignment = ["fill", "top"];
    var btnCol = ctrlGroup.add("button", undefined, "按列解析");
    btnCol.alignment = ["fill", "top"];


    var jsonGroup = ctrlGroup.add('group');
    jsonGroup.alignment = ["right", "top"];
    var jsonBtn = jsonGroup.add('button', undefined, 'v 展开JSON', { name: 'expand' });
    jsonBtn.alignment = ["right", "top"];
    autoSizeButton(jsonBtn);
    jsonBtn.onClick = function () {
        // resultBox.visible = !resultBox.visible;
        if (resultBox.size[1] > 2) {
            // resultBox.size = [resultBox.size[0], 2]; // 高度设为 2
            animateSize(win, resultBox, [resultBox.size[0], 2], 50, 10);
            this.text = 'v 展开JSON';
        } else {
            // resultBox.size = [resultBox.size[0], 100]; // 恢复高度
            // resultBox.size = [resultBox.size[0], 100]; // 恢复高度
            animateSize(win, resultBox, [resultBox.size[0], 100], 50, 10);

            this.text = '^ 收起JSON';
        }
        win.layout.layout(true);
    };



    // === 解析结果区（json） ===
    var resultBox = win.add("edittext", undefined, "", { multiline: true, scrolling: true });
    resultBox.preferredSize.height = 2;
    resultBox.alignment = ["fill", "top"];
    resultBox.margins = 0;

    // === 滚动展示区 ===

    var scrollBlock = win.add("panel", undefined, "");
    scrollBlock.size = [550, 280];
    scrollBlock.maximumSize.width = 550;
    scrollBlock.minimumSize.height = 280;
    scrollBlock.alignment = ["fill", "fill"];
    scrollBlock.alignChildren = ["fill", "fill"];
    scrollBlock.layout.layout(true);
    scrollBlock.margins = 0;

    var resultTitleShow = true;
    var resultTitle = scrollBlock.add("statictext", undefined, "解析后结果将在这里显示：");
    resultTitle.orientation = "stack";
    resultTitle.alignChildren = ["left", "top"];
    resultTitle.alignment = ["left", "top"];



    // ✅ 新增包裹容器（修复布局偏移问题）
    var scrollContainer = scrollBlock.add("group");
    scrollContainer.orientation = "column";
    scrollContainer.alignChildren = ["fill", "top"];
    scrollContainer.alignment = ["fill", "fill"];
    scrollContainer.margins = 0;

    // ✅ 滚动内容区
    var scrollGroup = scrollContainer.add("group");
    scrollGroup.orientation = "row";
    scrollGroup.alignChildren = ["left", "top"];
    scrollGroup.alignment = ["left", "top"];
    scrollGroup.margins = 10;
    scrollGroup.maximumSize.width = 100000000;
    scrollGroup.layout.layout(true);

    var scrollArea = scrollBlock.add("group");
    scrollArea.orientation = "stack";
    scrollArea.alignChildren = ["fill", "fill"];
    scrollArea.preferredSize = [580, 30];
    scrollArea.alignment = ["fill", "bottom"];
    scrollArea.margins = 0;
    scrollArea.layout.layout(true);

    // ✅ 横向滚动条
    var scrollBar = scrollArea.add("scrollbar");
    scrollBar.stepdelta = 10;
    scrollBar.alignment = ["fill", "bottom"];



    scrollBar.onChanging = function () {
        scrollGroup.location.x = -1 * this.value; // 横向滚动
    };

    // === 更新滚动条范围 ===
    function scrollBarUpdate() {
        scrollBar.size = [scrollBlock.size.width, 20]; // 横向 scrollbar
        scrollBar.maxvalue = Math.max(0, scrollGroup.size.width - scrollBlock.size.width + 10);
        scrollBar.value = 0; // ✅ 确保初始为 0
        scrollBar.alignment = ["fill", "bottom"];

        tipText.text = scrollGroup.children.length != 0 ? "在这滚动" : "填写文本后解析";
    }

    win.onShow = function () {
        win.layout.layout(true);
        scrollBarUpdate();
        win.layout.layout(true);
    };


    var tipGroup = scrollArea.add("group");
    tipGroup.preferredSize = [30, 30];
    tipGroup.orientation = "column";
    tipGroup.alignment = ["center", "fill"];
    var tipText = tipGroup.add("statictext", undefined, "填写文本后解析");
    tipText.graphics.font = ScriptUI.newFont("Arial", "Regular", 15); // 字体: Arial，粗体，10pt

    tipText.alignment = ["center", "center"];

    var g = tipText.graphics;
    var redPen = g.newPen(g.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 0.5); // RGB [1,0,0]，alpha 0.5
    g.foregroundColor = redPen;


    //   var btnNext = ctrlGroup.add("button", undefined, "下一步");
    //     autoSizeButton(btnNext);
    //     btnNext.onClick = function () {
    //         alert('\n'+resultBox.text);
    //     };


    // === 工具函数 ===
    function parseExcelText(text) {
        var lines = text.split(/\r?\n/);
        var result = [];
        var buffer = "";
        var inQuote = false;

        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            if (line.replace(/\s/g, "") === "") continue;

            // 如果前面是多行单元格，累加
            if (inQuote) {
                buffer += "\n" + line;
            } else {
                buffer = line;
            }

            // 统计当前行中引号的数量
            var quoteCount = 0;
            for (var j = 0; j < line.length; j++) {
                if (line.charAt(j) === '"') quoteCount++;
            }

            // 判断引号是否成对关闭
            if (quoteCount % 2 !== 0) {
                inQuote = !inQuote;
            }

            if (!inQuote) {
                // 一行完整数据
                result.push(buffer);
                buffer = "";
            }
        }

        // 拆分列（注意：要处理单元格内的 tab 分割）
        var table = [];
        for (var i = 0; i < result.length; i++) {
            var line = result[i];
            // 去掉包裹的引号，并恢复内部换行
            var cells = [];
            var cell = "";
            var inQ = false;
            for (var k = 0; k < line.length; k++) {
                var ch = line.charAt(k);
                if (ch === '"') {
                    inQ = !inQ;
                } else if (ch === "\t" && !inQ) {
                    cells.push(cell);
                    cell = "";
                } else {
                    cell += ch;
                }
            }
            cells.push(cell);
            table.push(cells);
        }

        return table;
    }

    function parseByRow(table) {
        if (table.length < 2) return table;
        var headers = table[0];
        var objs = [];
        for (var i = 1; i < table.length; i++) {
            var row = table[i];
            var obj = {};
            for (var j = 0; j < headers.length; j++) {
                obj[headers[j]] = row[j];
            }
            objs.push(obj);
        }
        return objs;
    }

    function parseByCol(table) {
        if (table.length === 0) return table;
        var headers = [];
        for (var i = 0; i < table.length; i++) {
            headers.push(table[i][0]);
        }
        var colCount = table[0].length;
        var objs = [];
        for (var j = 1; j < colCount; j++) {
            var obj = {};
            for (var i = 0; i < headers.length; i++) {
                obj[headers[i]] = table[i][j];
            }
            objs.push(obj);
        }
        return objs;
    }


    function detectFormat(table) {
        if (table.length == 0) return "row";
        var firstRowKeys = {};
        var firstColKeys = {};
        for (var i = 0; i < table[0].length; i++) firstRowKeys[table[0][i]] = true;
        for (var j = 0; j < table.length; j++) firstColKeys[table[j][0]] = true;

        var rowUniq = 0, colUniq = 0;
        for (var k in firstRowKeys) rowUniq++;
        for (var k2 in firstColKeys) colUniq++;
        return (colUniq > rowUniq && table.length > 1) ? "col" : "row";
    }

    function clearCards() {
        while (scrollGroup.children.length > 0) {
            scrollGroup.remove(scrollGroup.children[0]);
        }
    }

    function truncateString(str, maxLen) {
        if (!str) return "";
        str = String(str);
        if (str.length <= maxLen) return str;
        var headLen = Math.ceil(maxLen / 2);
        var tailLen = Math.floor(maxLen / 2);
        return str.substr(0, headLen) + "…" + str.substr(str.length - tailLen, tailLen);
    }

    // === 渲染卡片 ===
    function renderCards(dataArray) {
        clearCards();
        if (resultTitleShow == true) {
            resultTitle.parent.remove(resultTitle);
            resultTitleShow = false;
        }

        if (!dataArray || dataArray.length === 0) return;


        var cardsNum = 0;
        var maxLen = parseInt(lenInput.text, 10);
        if (isNaN(maxLen) || maxLen <= 0) maxLen = 15;

        var keys = [];
        for (var k in dataArray[0]) {
            keys.push(k);
        }

        var keyCard = scrollGroup.add("panel", undefined, "Keys");
        keyCard.orientation = "column";
        keyCard.alignChildren = ["fill", "top"];
        keyCard.margins = 8;
        keyCard.preferredSize.width = 150;
        for (var i = 0; i < keys.length; i++) {
            keyCard.add("statictext", undefined, truncateString(keys[i], maxLen));
        }

        var card;
        for (var idx = 0; idx < dataArray.length; idx++) {
            card = scrollGroup.add("panel", undefined, "Item " + (idx + 1));
            card.orientation = "column";
            card.alignChildren = ["fill", "top"];
            card.margins = 8;
            card.preferredSize.width = 150;
            var obj = dataArray[idx];
            for (var j = 0; j < keys.length; j++) {
                var val = obj[keys[j]];
                card.add("statictext", undefined, truncateString(val, maxLen));
            }
            cardsNum++;
        }


        win.layout.layout(true);
        scrollBarUpdate();
        win.layout.layout(true);

        var groupHeight = card.size[1] + card.margins * 2;
        scrollGroup.preferredSize.height = groupHeight;

        win.layout.layout(true);
    }

    // === 按钮事件 ===
    btnRow.onClick = function () {
        var table = parseExcelText(inputBox.text);
        var parsed = parseByRow(table);
        thisData.currentResult = parsed;
        renderCards(parsed);
        resultBox.text = toJSONString(parsed, ' ', 2);
    };

    btnCol.onClick = function () {
        var table = parseExcelText(inputBox.text);
        var parsed = parseByCol(table);
        thisData.currentResult = parsed;
        renderCards(parsed);
        resultBox.text = toJSONString(parsed, ' ', 2);
    };

    btnAuto.onClick = function () {
        var table = parseExcelText(inputBox.text);
        var mode = detectFormat(table);
        var parsed = (mode == "col") ? parseByCol(table) : parseByRow(table);
        thisData.currentResult = parsed;
        renderCards(parsed);
        resultBox.text = toJSONString(parsed, ' ', 2);
        alert("自动检测结果：按 " + (mode == "col" ? "列" : "行") + " 解析。");
    };

    // var that = this; 
    // if(fnsObj){
    //     var btnGroup = win.add("group");
    //     btnGroup.orientation = "row";
    //     btnGroup.alignChildren = ["fill", "top"];
    //     for (var label in fnsObj) {
    //         if (fnsObj.hasOwnProperty(label)) {
    //             var btn = btnGroup.add("button", undefined, label);

    //             // 可选：给按钮附加自定义数据
    //             btn.data = this.data;

    //             // 使用闭包将当前 btn 传给回调函数
    //             (function(button, actionFn) {
    //                 button.onClick = function() {
    //                     actionFn.call(this, button.data); // 或直接传 button 本身
    //                 };
    //             })(btn, fnsObj[label]);
    //         }
    //     }
    // }


    // === 动态按钮区 ===
    if (fnsDict) {
        var btnGroup = win.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignChildren = ["fill", "top"];

        for (var label in fnsDict) {
            if (fnsDict.hasOwnProperty(label)) {

                var btn = btnGroup.add("button", undefined, label);

                (function (actionFn, dataRef) {
                    btn.onClick = function () {
                        // actionFn(dataRef.currentResult); // 或 actionFn.call(null, dataRef)
                        win.close();
                        actionFn(dataRef.currentResult);

                    };
                })(fnsDict[label], thisData);

            };
        };
    };



    win.show();

}



