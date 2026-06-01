/**
 * Select objects in Illustrator by name and type, within a given scope.
 * @param {Object} scope - Document, GroupItem, Layer, Array, or any PageItem container.
 * @param {String} name - The name pattern to match (supports * as wildcard).
 * @param {String} [type] - The type of object (e.g., 'group', 'text', 'path').
 * @param {Boolean} [exactMatch] - If true, name must match exactly.
 */
function querySelect(scope, name, type, exactMatch) {
    // Type mapping for common abbreviations
    var typeMap = {
        'group'         : 'GroupItem',
        'grp'           : 'GroupItem',
        'text'          : 'TextFrame',
        'txt'           : 'TextFrame',
        'path'          : 'PathItem',
        'paths'         : 'PathItem',
        'compound'      : 'CompoundPathItem',
        'compoundpath'  : 'CompoundPathItem',
        'compoundpathitem': 'CompoundPathItem',
        'img'           : 'PlacedItem',
        'image'         : 'PlacedItem',
        'placed'        : 'PlacedItem',
        'embed'         : 'RasterItem',
        'embedded'      : 'RasterItem',
        'raster'        : 'RasterItem',
        'rasteritem'    : 'RasterItem',
        'symbol'        : 'SymbolItem',
        'mesh'          : 'MeshItem',
        'plugin'        : 'PluginItem',
        'graph'         : 'GraphItem',
        'legacytext'    : 'LegacyTextItem',
        'clippingmask'  : 'PathItem',           // masks are PathItems with .clipping = true
        'area'          : 'TextFrame',
        'point'         : 'TextFrame'
    };

    // Parse [artboard]name<type> or [artboard]name
    var artboardPattern = /^\[([^\]]+)\](.+)$/;
    var typePattern = /^(.+)<(.+)>$/;
    var artboardName = null;
    var namePattern = name;
    var aiType = null;

    // Extract artboard name if present
    var m = name.match(artboardPattern);
    if (m) {
        artboardName = m[1];
        namePattern = m[2];
    }

    // Extract type if present in namePattern
    if (!type && typePattern.test(namePattern)) {
        var t = namePattern.match(typePattern);
        namePattern = t[1];
        type = t[2];
    }

    // Normalize type
    if (type) {
        type = type.toLowerCase();
        aiType = typeMap[type] || type;
    }

    // Helper: match name with wildcards
    function matchName(objName, pattern, exact) {
        if (exact) return objName == pattern;
        var re = new RegExp('^' + pattern.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, '\\$1').replace(/\\\*/g, '.*') + '$', 'i');
        return re.test(objName);
    }

    // Helper: get all items from scope
    function getItemsFromScope(scopeObj) {
        if (!scopeObj) return [];
        if (scopeObj.typename == "Document") return scopeObj.pageItems;
        if (scopeObj.typename == "GroupItem" || scopeObj.typename == "Layer") return scopeObj.pageItems;
        if (scopeObj.length !== undefined && typeof scopeObj.length === "number") return scopeObj; // Array or selection
        if (scopeObj.typename && scopeObj.pageItems) return scopeObj.pageItems;
        return [];
    }

    var doc = app.activeDocument;
    if (!doc) return;

    // Deselect all first
    doc.selection = null;

    var items = getItemsFromScope(scope);
    if (!items || items.length == 0) return [];

    // Filter by artboard if specified
    var artboardIdx = null;
    if (artboardName) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (doc.artboards[i].name == artboardName) {
                artboardIdx = i;
                break;
            }
        }
    }

    var found = [];
    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        // Artboard filter
        if (artboardName !== null) {
            var abIdx = doc.artboards.getActiveArtboardIndex();
            // Use geometricBounds to check if item is on artboard
            var abRect = doc.artboards[artboardIdx].artboardRect;
            var gb = item.geometricBounds;
            // Check if item's center is within artboard bounds
            var cx = (gb[0] + gb[2]) / 2;
            var cy = (gb[1] + gb[3]) / 2;
            if (!(cx >= abRect[0] && cx <= abRect[2] && cy <= abRect[1] && cy >= abRect[3])) continue;
        }

        // Type filter
        if (aiType && item.typename != aiType) continue;

        // Name filter
        if (matchName(item.name, namePattern, exactMatch)) {
            found.push(item);
        }
    }
    if (found.length) {
        doc.selection = found;
    }
    return found;
}

// 用法示例：
querySelect(app.activeDocument, "[ab1]*op*<txt>"); 
// querySelect(app.activeDocument, "[Artboard1]*logo*<text>");
// querySelect(app.activeDocument, "*button*");
// querySelect(myGroup, "Header", "group", true);
// querySelect(app.activeDocument, "[Artboard2]*", "path");
