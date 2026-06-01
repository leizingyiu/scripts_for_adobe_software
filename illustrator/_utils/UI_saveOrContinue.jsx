function UI_saveOrContinue() {
    if (app.documents.length === 0) {
        alert("没有打开的文档！");
        return false;
    }

    var doc = app.activeDocument;

    if (doc.saved) {
        return true;
    }

    var dialog = new Window("dialog", "保存确认");
    dialog.orientation = "column";

    var textGroup = dialog.add("group");
    textGroup.add("statictext", undefined, "文件 '" + doc.name + "' 已被修改。您想保存更改吗？");

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";

    // 按钮的返回值：1 = 保存并继续, 2 = 不保存并继续, 3 = 取消 (退出)

    var saveBtn = buttonGroup.add("button", undefined, "保存并继续", { name: "save" });
    var discardBtn = buttonGroup.add("button", undefined, "不保存并继续", { name: "discard" });
    var cancelBtn = buttonGroup.add("button", undefined, "取消", { name: "cancel" });

    saveBtn.onClick = function () {
        dialog.close(1); // 返回 1
    };

    discardBtn.onClick = function () {
        dialog.close(2); // 返回 2
    };

    cancelBtn.onClick = function () {
        dialog.close(3); // 返回 3
    };

    // 设置默认按钮（按下 Enter 键）和取消按钮（按下 Esc 键）
    dialog.defaultElement = saveBtn;
    dialog.cancelElement = cancelBtn;

    var response = dialog.show(); // 显示对话框并等待用户响应

    if (response === 1) {
        // 保存并继续
        try {
            doc.save();
            return true;
        } catch (e) {
            throw new Error("保存操作被取消或发生错误: \n" + e + '\n _main_getAndRun.js');
            return false; // 保存失败，退出脚本
        }
    } else if (response === 2) {
        // 不保存并继续
        return true;
    } else { // response === 3 或窗口被关闭
        // 取消（退出）
        return false;
    }
}


// var docSaved = UI_saveOrContinue();
// if (!docSaved) {
//     alert("用户已取消");
// }else{
//     alert("用户已保存"); 
//     main();
// }