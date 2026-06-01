// ============================================================
// Illustrator Script (ES3) by le
// Text Replace Tool (Preserve Styles + Post-check Mixed Styles)
// ============================================================

(function () {

    // --- 获取目标文本项 ---
    function getTargetTextItems() {
        var doc = app.activeDocument;
        var items = [];

        if (app.selection.length > 0) {
            for (var i = 0; i < app.selection.length; i++) {
                if (app.selection[i].typename === "TextFrame") {
                    items.push(app.selection[i]);
                }
            }
        } else {
            if (confirm("没有选中项，是否处理所有文本对象？")) {
                var all = doc.textFrames;
                for (var j = 0; j < all.length; j++) {
                    items.push(all[j]);
                }
            }
        }
        return items;
    }

    // --- 获取样式段数量 ---
    function getStyleSegmentCount(tf) {
        try {
            var tr = tf.textRange;
            if (tr.textRanges && tr.textRanges.length)
                return tr.textRanges.length;
        } catch (e) {}
        return 1;
    }

    // --- 替换逻辑 ---
    function doReplace(items, searchStr, replaceStr, useRegex) {
        var changedStyleItems = [];
        var re;

        if (useRegex) {
            try {
                re = new RegExp(searchStr, "g");
            } catch (e) {
                alert("正则表达式错误：" + e);
                return;
            }
        }

        for (var i = 0; i < items.length; i++) {
            var tf = items[i];
            var oldText = tf.contents;

            if (!oldText || oldText === "") continue;

            // 替换前样式段数量
            var beforeCount = getStyleSegmentCount(tf);

            // 执行替换
            var newText = useRegex ? oldText.replace(re, replaceStr) : oldText.split(searchStr).join(replaceStr);

            if (newText !== oldText) {
                tf.contents = newText;
            } else {
                continue;
            }

            // 替换后样式段数量
            var afterCount = getStyleSegmentCount(tf);

            if (afterCount > beforeCount + 1) {
                changedStyleItems.push(tf);
            }
        }

        // 替换完成提示
        if (changedStyleItems.length > 0) {
            if (confirm("替换完成，但有 " + changedStyleItems.length + " 个文本对象替换后外观或样式差异较大，是否选中它们？")) {
                app.selection = changedStyleItems;
            }
        } else {
            alert("替换完成！");
        }
    }

    // --- 创建 UI ---
    function showDialog() {
        var dlg = new Window("dialog", "文本替换工具");
        dlg.alignChildren = "fill";

        var useRegexGroup = dlg.add("group");
        var useRegexCheck = useRegexGroup.add("checkbox", undefined, "使用正则表达式");

        var inputGroup1 = dlg.add("group");
        inputGroup1.add("statictext", undefined, "搜索内容：");
        var searchInput = inputGroup1.add("edittext", undefined, "");
        searchInput.characters = 30;

        var inputGroup2 = dlg.add("group");
        inputGroup2.add("statictext", undefined, "替换内容：");
        var replaceInput = inputGroup2.add("edittext", undefined, "");
        replaceInput.characters = 30;

        var btnGroup = dlg.add("group");
        btnGroup.alignment = "center";

        var runBtn = btnGroup.add("button", undefined, "执行替换");
        var closeBtn = btnGroup.add("button", undefined, "关闭");

        var items = getTargetTextItems();
        if (items.length === 0) {
            alert("没有可处理的文本对象。");
            return;
        }

        runBtn.onClick = function () {
            var searchStr = searchInput.text;
            var replaceStr = replaceInput.text;
            var useRegex = useRegexCheck.value;

            if (searchStr === "") {
                alert("请输入搜索内容。");
                return;
            }

            doReplace(items, searchStr, replaceStr, useRegex);
            app.redraw();
        };

        closeBtn.onClick = function () {
            dlg.close();
        };

        dlg.show();
    }

    showDialog();

})();
