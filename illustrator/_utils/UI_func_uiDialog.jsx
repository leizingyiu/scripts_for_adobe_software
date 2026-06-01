/**
 * 通用左对齐对话框（自动识别标题）
 * 第一行 = 加粗标题；其余 = 正文
 *
 * @param {"alert"|"confirm"|"prompt"} type - 对话框类型
 * @param {string} message - 显示内容，第一行会作为标题
 * @param {object} [options] - 可选项
 * @param {number} [options.width=400] - 初始窗口宽度
 * @param {string} [options.defaultText=""] - prompt 默认值
 * @returns {boolean|string|undefined}
 */
function uiDialog(type, message, options) {
    options = options || {};
    var initialWidth = options.width || 600;
    var defaultText = options.defaultText || "";

    // 拆分 message：第一行为标题，其他为正文
    var lines = message.split(/\r?\n/);
    var titleLine = lines.shift() || "";
    var bodyText = lines.join("\n");

    // var win = new Window("dialog", " ", undefined, { resizeable: true });
    var win = new Window("dialog", " ", undefined);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"]; // 子元素宽度撑满，垂直从上开始
    win.margins = 20;

    // ====== 标题 ======
    var title;
    if (titleLine) {
        title = win.add("statictext", undefined, titleLine);
        title.justify = "left";
        title.graphics.font = ScriptUI.newFont("dialog", "bold", 16);
        // 不设 preferredSize.width，靠 alignChildren 自动撑开
    }

    // ====== 正文（可滚动、只读、撑开区域） ======
    var text;
    if (bodyText) {
        text = win.add("edittext", undefined, bodyText, { multiline: true, scrolling: true });
        // text.enabled = false; // 禁止编辑（更安全）
        text.active = false;
        // 关键：不设 height，让其自动伸缩；宽度由父容器控制
        text.alignment = ["fill", "fill"];
        // 设置最小高度，避免太小
        text.minimumSize = [200, 80];
        // 允许垂直拉伸（通过布局，不是 preferredSize）
        text.alignment = ["fill", "fill"];
    }

    // ====== 输入框（仅 prompt） ======
    var input;
    if (type === "prompt") {
        input = win.add("edittext", undefined, defaultText);
        input.alignment = ["fill", "bottom"];
        input.minimumSize = [0, 25];
    }

    // ====== 按钮组 ======
    var group = win.add("group");
     group.alignment = "center", group.orientation = "row", group.alignChildren = ["center", "center"];

    var okBtn = group.add("button", undefined, "确定", { name: "ok" });
    var cancelBtn;
    var result;

    if (type === "alert") {
        okBtn.onClick = function () { result = undefined; win.close(); };
    } else if (type === "confirm") {
        cancelBtn = group.add("button", undefined, "取消", { name: "cancel" });
        okBtn.onClick = function () { result = true; win.close(); };
        cancelBtn.onClick = function () { result = false; win.close(); };
    } else if (type === "prompt") {
        cancelBtn = group.add("button", undefined, "取消", { name: "cancel" });
        okBtn.onClick = function () { result = input.text; win.close(); };
        cancelBtn.onClick = function () { result = null; win.close(); };
    }

    // ====== 关键：设置窗口初始大小，并让布局生效 ======
    // 设置窗口最小尺寸
    win.minimumSize = [300, 200];
    // 初始大小（宽度由参数控制，高度自适应）

    // 强制布局一次
    win.layout.layout(true);

    win.show();
    win.size = [initialWidth, win.size[1]]; // 保持当前高度，只设宽度
    win.layout.layout(true);
    
    return result;
}