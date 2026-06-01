

function setVisible(item, vis) {
    switch (true) {
        case item.typename == "Layer":
            item.visible = Boolean(vis);
            break;
        case item.typename == "Document":
            throw new Error("请选择图层或pageitem - setVisible() - func_visible.jsx");
            break;
        default:
            if(typeof(item.hidden)!="undefined"){
                item.hidden = !Boolean(vis);
            }else{
                throw new Error("请选择图层或pageitem - setVisible() - func_visible.jsx");
            }
    }
}

function getVisible(item){
    switch (true) {
        case item.typename == "Layer":
            return item.visible;
            break;
        case item.typename == "Document":
            throw new Error("请选择图层或pageitem - getVisible() - func_visible.jsx");
            break;
        default:
            if(typeof(item.hidden)!="undefined"){
                return !item.hidden;
            }else{
                throw new Error("请选择图层或pageitem - getVisible() - func_visible.jsx");
            }
    }
}



function visible_testing() {

 
    var doc = app.activeDocument;
    if (!doc) {
        alert("请先打开一个 Illustrator 文档。");
    }

    setVisible(doc.layers[0], false);
    app.redraw();
    alert("第 0 个图层已隐藏 - 1/4");

    setVisible(doc.layers[0], true);
    app.redraw();
    alert("第 0 个图层已显示 - 2/4");

    setVisible(doc.pageItems[0], false);
    app.redraw();
    alert("第 0 对象已隐藏 - 3/4");

    setVisible(doc.pageItems[0], true);
    app.redraw();
    alert("第 0 对象已显示 - 4/4");

}

// visible_testing();

