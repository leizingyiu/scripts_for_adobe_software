#include "_utils/utils.jsx"

(function () {

    if (app.documents.length === 0) {
        alert("没有打开的文档。");
        return;
    }

    var doc = app.activeDocument;
    var i;



    // === Step 1. 另存为 source.ai ===
    var file = doc.fullName;
    var path = file.path;
    var name = file.name.replace(/\.ai$/i, '');
    var sourceFile = new File(path + "/" + name + "_source.ai");

    // 另存一份 AI 文件
    var saveOpts = new IllustratorSaveOptions();
    saveOpts.compatibility = Compatibility.ILLUSTRATOR17; // CS6+
    saveOpts.pdfCompatible = true;
    saveOpts.embedLinkedFiles = true; // 新增：嵌入链接的图片
    doc.saveAs(sourceFile, saveOpts);




    // ---  遍历 pageItems: 解锁 + 删除隐藏项 ---
    var allItems = [];
    for (i = 0; i < doc.pageItems.length; i++) {
        allItems.push(doc.pageItems[i]);
    }

    for (i = 0; i < allItems.length; i++) {
        var item = allItems[i];
        try {
            // 解锁
            if (item.locked) {
                item.locked = false;
            }

            // 删除隐藏对象
            if (item.hidden) {
                item.remove();
            }
        } catch (e) {
            // 忽略单个对象异常
        }
    }

    app.redraw();

    // 嵌入所有链接图片
    var linkedFiles = doc.placedItems;
    for (i = linkedFiles.length - 1; i >= 0; i--) {
        linkedFiles[i].embed();
    }

    app.redraw();


    // ---  全选 + 扩展 + 全选 + 扩展外观 ---
    try {
        
        app.selection = null;
        app.executeMenuCommand("selectall");
        app.executeMenuCommand("expandStyle");  // 扩展外观
        app.redraw();

        app.selection = null;
        app.executeMenuCommand("selectall");
        app.executeMenuCommand("Expand3");       // 扩展
        app.redraw();

    } catch (e) {
        alert("扩展命令执行失败：" + e + ' : ' + e.line);
    }

    // --- 3️⃣ 导出 PDF ---
    try {
        var file = doc.fullName;
        var path = file.path;
        var name = file.name.replace(/\.ai$/i, '');

        name = name.replace('source', 'export');

        var pdfFile = new File(path + "/" + name + ".pdf");
        var pdfOpts = new PDFSaveOptions();
        pdfOpts.compatibility = PDFCompatibility.ACROBAT4;
        pdfOpts.preserveEditability = false;

        doc.saveAs(pdfFile, pdfOpts);

        alert("✅ 完成：隐藏项已删除，所有对象已扩展，并导出 PDF，打开文件夹看看吧。");

        revealFolder();
        
        doc.close();

    } catch (e) {
        alert("导出 PDF 失败：" + e);
    }
})();
