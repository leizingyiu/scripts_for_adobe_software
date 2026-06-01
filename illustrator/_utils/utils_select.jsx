#include "utils_artboard.jsx"

/**
 * 选取 Illustrator 中的一个或多个页面项 (PageItem)。
 * 该函数会清除当前选区，然后将传入的所有有效项目设置为选中状态。
 *
 * 兼容以下调用方式：
 * 1. select(item) - 传入单个项目。
 * 2. select([item1, item2, item3...]) - 传入项目数组。
 * 3. select(item1, item2, item3...) - 传入多个项目作为参数。
 *
 * @param {PageItem|PageItem[]|...PageItem} items 一个或多个 PageItem 对象。

 return : 选中的项目数量。
 */
function select() {

    if (app.documents.length === 0) {
        throw new Error("No active document found.");
        return;
    }

    var doc = doc || app.activeDocument,
        itemsToSelect = [];

    if (arguments.length === 1 && arguments[0]instanceof Array) {
        itemsToSelect = arguments[0];
    } else {
        for (var i = 0; i < arguments.length; i++) {
            itemsToSelect.push(arguments[i]);
        }
    }

    doc.selection = null;

    var selectedCount = 0;

    var _selection = [];

    var collectLayerItems = function (layer, result) {
        if (!(layer && layer.pageItems)) 
            return;
        for (var i = 0; i < layer.pageItems.length; i++) {
            result.push(layer.pageItems[i]);
        }
        for (var j = 0; j < layer.layers.length; j++) {
            collectLayerItems(layer.layers[j], result);
        }
    };

    for (var j = 0; j < itemsToSelect.length; j++) {
        var item = itemsToSelect[j];
        if (item && item.typename) {

            var type = item.typename;

            // if (type === "PathItem" || type === "TextFrame" || type === "GroupItem" || type === "CompoundPathItem" || type === "RasterItem" || type === "MeshItem" || type === "SymbolItem" || type === "PlacedItem" || type === "LegacyTextItem" || type === "GraphItem") {
            //     try {
            //         _selection.push(item);
            //     } catch (e) {
            //         throw new Error("Failed to select item " + item.name + ": " + e.message);
            //     }
            // } else if (type === "Layer") {
            //     try {
            //         collectLayerItems(item, _selection);
            //     } catch (e) {
            //         throw new Error("Failed to select items in layer " + item.name + ": " + e.message);
            //     }
            // }

            switch (type) {
                case "PathItem":
                case "TextFrame":
                case "GroupItem":
                case "CompoundPathItem":
                case "RasterItem":
                case "MeshItem":
                case "SymbolItem":
                case "PlacedItem":
                case "LegacyTextItem":
                case "GraphItem":
                    try {
                        _selection.push(item);
                    } catch (e) {
                        throw new Error("Failed to select item " + item.name + ": " + e.message);
                    }
                    break;
                    
                case "Layer":
                    try {
                        collectLayerItems(item, _selection);
                    } catch (e) {
                        throw new Error("Failed to select items in layer " + item.name + ": " + e.message);
                    }
                    break;
                
                case "Artboard":                    
                    processArtboardsByLayers(item, function (items) {
                        for (var i = 0; i < items.length; i++) {
                            _selection.push(items[i]);
                        }
                    });

                default:
                    throw new Error("Unsupported item type: " + type+' /n select() utils_select.jsx');
                    break;
            }
        }
    }

    try {
        doc.selection = _selection;
        selectedCount = doc.selection.length;
    } catch (e) {

        selectedCount = 0;
        for (var j = 0; j < itemsToSelect.length; j++) {
            var item = itemsToSelect[j];
            if (item && item.typename) {
                var type = item.typename;
                if (type === "PathItem" || type === "TextFrame" || type === "GroupItem" || type === "CompoundPathItem" || type === "RasterItem" || type === "MeshItem" || type === "SymbolItem" || type === "PlacedItem" || type === "LegacyTextItem" || type === "GraphItem") {
                    try {
                        item.selected = true;
                        selectedCount++;
                    } catch (e) {
                        throw new Error("Failed to select item" + item.name + ": " + e.message);
                    }
                } else if (type === "Layer") {
                    try {
                        var layerQueue = [item];

                        while (layerQueue.length > 0) {
                            var currentLayer = layerQueue.shift();

                            for (var i = 0; i < currentLayer.pageItems.length; i++) {
                                currentLayer.pageItems[i].selected = true;
                                selectedCount++;
                            }

                            for (var j = 0; j < currentLayer.layers.length; j++) {
                                layerQueue.push(currentLayer.layers[j]);
                            }
                        }
                    } catch (e) {
                        throw new Error("Failed to select items in layer " + item.name + ": " + e.message);
                    }
                }else if(type === "Artboard"){
                    var __sel = doc.selection; 

                    doc.artboards.setActiveArtboardIndex(indexOfArtboard(item));
                    app.executeMenuCommand('selectallinartboard');
                    var artboardSelection = doc.selection;

                    var newSelection = [];
                    if (__sel && __sel.length > 0) {
                        for (var i = 0; i < __sel.length; i++) {
                            newSelection.push(__sel[i]);
                        }
                    }
                    if (artboardSelection && artboardSelection.length > 0) {
                        for (var j = 0; j < artboardSelection.length; j++) {
                            newSelection.push(artboardSelection[j]);
                        }
                    }

                    doc.selection = newSelection;
                }
            }
        }

    }

    return selectedCount;

}

function select_testing() {
    var doc = app.activeDocument;
    // var item = doc.pageItems[0]; var item2 = doc.pageItems[1]; var item3 =
    // doc.pageItems[2]; var count = select(item, item2, item3); alert("选中 " + count
    // + " 个项目。");

    for (var i = 0; i < doc.pageItems.length; i++) {
        select(doc.pageItems[i]);
        app.redraw();
        alert("第 " + (i + 1) + " 个项目 " + doc.pageItems[i].name + " 已选中 - " + (i + 1) + '/' + doc.pageItems.length);
    }
    for (var i = 0; i < doc.layers.length; i++) {
        select(doc.layers[i]);
        app.redraw();
        alert("第 " + (i + 1) + " 个图层 " + doc.layers[i].name + " 已选中 - " + (i + 1) + "/" + doc.layers.length);
    }

    select(doc.artboards[0]);

}

// select_testing();
