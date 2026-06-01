#target illustrator





function getSingleImageFromGroup(item) {
    var fnAndfile = "/n/n getSingleImageFromGroup() - imgReplace.jsx";
    if (item.typename !== "GroupItem") {
        throw new Error("Input item is not a GroupItem." + fnAndfile);
    }

    var allImageItems = [];

    // 递归收集函数
    function collectImages(pageItem) {
        if (!pageItem || !pageItem.pageItems) {
            return;
        }
        var children = pageItem.pageItems;
        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            if (child.typename === "RasterItem" || child.typename === "PlacedItem") {
                allImageItems.push(child);
            } else if (child.typename === "GroupItem") {
                collectImages(child); // 递归进入子组
            }
            // 忽略其他类型（PathItem, TextFrame 等）
        }
    }

    // 开始递归收集
    collectImages(item);

    if (allImageItems.length === 0) {
        alert("Error: No image found in the group (including nested groups).");
        throw new Error("No image found in the group." + fnAndfile);
    } else if (allImageItems.length > 1) {
        alert("Error: More than one image found in the group (including nested groups).");
        throw new Error("More than one image found in the group." + fnAndfile);
    }

    return allImageItems[0];
}

function findImageFile(docFolder, imgname) {
    var folder = new Folder(docFolder);
    if (!folder.exists) {
        throw new Error("文件夹不存在: " + docFolder);
    }

    if (imgname === "" || !imgname) {
        throw new Error("请输入图片名称。");
    }
    var imgFile = new File(folder.fsName + "/" + imgname);

    // 如果传入的 imgname 已带后缀且存在
    if (imgFile.exists) {
        return imgFile;
    }

    // 去掉已有的扩展名部分（兼容用户传入 "logo.png"）
    var baseName = imgname.replace(/\.[^\.]+$/, "");

    // 常见图片后缀名白名单（可自行扩展）
    var allowedExts = [
        ".png",
        ".jpg",
        ".jpeg",
        ".tif",
        ".tiff",
        ".bmp",
        ".gif",
        ".psd",
        ".svg",
        ".webp"
    ];

    //  // 在文件夹中查找符合条件的文件 var matched = folder.getFiles(function (f) {     if (!(f
    // instanceof File)) return false;     var nameNoExt =
    // f.name.replace(/\.[^\.]+$/, "");     var ext = f.name.match(/\.[^\.]+$/);
    // if (!ext) return false;     var extLower = ext[0].toLowerCase();     return
    // nameNoExt === baseName && allowedExts.indexOf(extLower) !== -1; });
    // 用于记录所有被测试的文件名
    var testedFiles = [];

    var matched = folder.getFiles(function (f) {
        // 记录当前正在测试的文件（只记录文件，不记录文件夹）
        if (f instanceof File) {
            testedFiles.push(f.name);
        } else {
            // 如果是文件夹，也可以记录（可选） testedFiles.push("[Folder] " + f.name);
            return false;
        }

        var nameNoExt = f
            .name
            .replace(/\.[^\.]+$/, "");
        var ext = f
            .name
            .match(/\.[^\.]+$/);
        if (!ext) 
            return false;
        var extLower = ext[0].toLowerCase();
        return nameNoExt === baseName && allowedExts.indexOf(extLower) !== -1;
    });

    // // 运行完后弹出所有被测试的文件 if (testedFiles.length > 0) {     alert("Tested files:\n" +
    // testedFiles.join("\n")); } else {     alert("No files were tested."); }

    if (matched.length === 0) {
        throw new Error("找不到匹配的图片文件: " + baseName);
    } else if (matched.length > 1) {
        var list = [];
        for (var i = 0; i < matched.length; i++) 
            list.push(matched[i].name);
        throw new Error("找到多个同名图片文件，请确认唯一性:\n" + list.join("\n"));
    }

    return matched[0];
}

function replaceImageByName(targetImageItemName, replacementImageFileName) { // example usage: replaceImageByName("old_image", "new_image.jpg")


  


    var targetImageItemName = targetImageItemName || "old_image",
        replacementImageFileName = replacementImageFileName || "new_image.jpg",
        doc = doc || app.activeDocument;

    if (!doc) {
        throw("请先打开一个 Illustrator 文档。");
        return;
    }

    var docFolder = doc.path,
        replacementImageFile = new File(docFolder + "/" + replacementImageFileName);

    // if (!replacementImageFile.exists) {     alert("替换图片文件不存在: " +
    // replacementImageFile.fsName);     return; }

    var foundCount = 0,
        replacedCount = 0;

    for (var i = doc.placedItems.length - 1; i >= 0; i--) {
        var item = doc.placedItems[i];
        if (item.name === targetImageItemName) {
            replaceImageCover(item, replacementImageFileName);
            foundCount++;
            replacedCount++;
        }
    }

    for (var j = doc.rasterItems.length - 1; j >= 0; j--) {
        var item = doc.rasterItems[j];
        if (item.name === targetImageItemName) {
            replaceImageCover(item, replacementImageFileName);
            foundCount++;
            replacedCount++;
        }
    }

    if (foundCount === 0) {
        alert("未找到名称为 '" + targetImageItemName + "' 的图像。");
    } else {
        alert("成功替换 " + replacedCount + " 个图像。");
    }
}

function replaceImageCover(oldItem, replaceSameFolderImageFileName) {
    log("replaceImageCover start");

    var isEmptyToContinue= isEmptyContent(replaceSameFolderImageFileName);
    log('isEmptyToContinue:' + isEmptyToContinue);

    if(isEmptyToContinue){
        if(log){
            log("replaceImageByName : "+ oldItem + " -> " + replaceSameFolderImageFileName +
            " : isEmptyContent(replacementImageFileName):"+isEmptyToContinue )
        }
        return false ; 
    }

    if (typeof oldItem[0] !== "undefined" && typeof oldItem[0].typename != "undefined") {

        for (var i = 0; i < oldItem.length; i++) {
            if (typeof oldItem[i].typename != "undefined") {
                replaceImageCover(oldItem[i], replaceSameFolderImageFileName);
            } else {
                throw new Error("替换图片时，似乎得到一个既非图形对象，也非数组的对象。 @replaceImageCover - imgReplace.jsx");
            }
        }
        return;
    }
    log("replaceImageCover oldItem: " + oldItem.name + " \\  replaceSameFolderImageFileName: " + replaceSameFolderImageFileName);

    var doc = doc || app.activeDocument,
        docFolder = doc.path,
        newImageFile = findImageFile(docFolder, replaceSameFolderImageFileName);

    log("replaceImageCover oldItem: " + oldItem.name + " \\  newImageFile: " + newImageFile.fsName );

    if (oldItem.typename === "GroupItem") {
        oldItem = getSingleImageFromGroup(oldItem);
    }

    var oldX = oldItem.position[0],
        oldY = oldItem.position[1],
        oldW = oldItem.width,
        oldH = oldItem.height,
        opacity = oldItem.opacity,
        blendMode = oldItem.blendingMode,
        layer = oldItem.layer,
        name = oldItem.name;

    // 获取原始尺寸
    var tempItem = layer
        .placedItems
        .add();
    tempItem.file = newImageFile;
    var originalW = tempItem.width,
        originalH = tempItem.height;
    tempItem.remove();


    


    var scaleX = oldW / originalW,
        scaleY = oldH / originalH,
        scale = Math.max(scaleX, scaleY),
        newW = originalW * scale,
        newH = originalH * scale;

    // 计算中心对齐位置（Y轴向上为正）
    var oldCenterX = oldX + oldW / 2;
    var oldCenterY = oldY - oldH / 2;
    var newLeft = oldCenterX - newW / 2;
    var newTop = oldCenterY + newH / 2;

    // 检查父级是否是剪切组
    var parentGroup = oldItem.parent;
    var isClippingGroup = false;
    var shouldCreateGroup = false;

    // alert([parentGroup.name, parentGroup.typename, parentGroup.clipped,
    // parentGroup.pageItems.length, oldItem.name, oldItem.typename].join('\n'))

    if (parentGroup && parentGroup.typename === "GroupItem" && parentGroup.clipped) {
        isClippingGroup = true;
    } else if ((parentGroup && parentGroup.typename === "GroupItem" && !parentGroup.clipped && parentGroup.pageItems.length > 1) || (parentGroup.typename === "Layer")) {
        shouldCreateGroup = true;
    }

   

    var finalGroup;
    if (shouldCreateGroup) {
        var finalGroup = parentGroup
            .groupItems
            .add();
        finalGroup.name = name;
        oldItem.move(finalGroup, ElementPlacement.INSIDE);

        // 更新变量，让后续逻辑把它当作“剪切组”处理
        parentGroup = finalGroup;

    }

    var newItem;

    

    if (isClippingGroup) {
        
       

        doc.selection=null; 
        // ✅ 是剪切组 → 放入组内
        newItem = parentGroup.placedItems.add(); // 看起来是这里闪退了，这个地方的add连 typeof 都不让

         
        newItem.file = newImageFile;
        newItem.width = newW;
        newItem.height = newH;
        newItem.position = [newLeft, newTop];

         newItem.opacity = opacity;
        newItem.blendingMode = blendMode;
        newItem.name = name;

         // 移动到组内最后（避免遮挡）
        newItem.move(parentGroup, ElementPlacement.PLACEATEND);

         oldItem.remove();

    } else {
         // ❌ 不是剪切组 → 创建剪切蒙版矩形
        newItem = layer
            .placedItems
            .add();
        newItem.file = newImageFile;
        newItem.width = newW;
        newItem.height = newH;
        newItem.position = [newLeft, newTop];

         // 创建剪切蒙版矩形（和原图尺寸相同，位置也相同）
        var clipRect = layer
            .pathItems
            .rectangle(oldY, // top
                    oldX, // left
                    oldW, // width
                    oldH, // height
            );
        var noColor = new NoColor();
        clipRect.fillColor = noColor; // 无填充
        clipRect.strokeColor = noColor; // 无描边

         // 把新图放到剪切蒙版下方（Z轴顺序） newItem.move(clipRect, ElementPlacement.PLACEBEFORE);

        clipRect.move(oldItem, ElementPlacement.PLACEBEFORE);
        newItem.move(oldItem, ElementPlacement.PLACEBEFORE);
        clipRect.clipping = true; // 设为剪切蒙版

         clipRect.position = [oldX, oldY];
        newItem.position = [newLeft, newTop];

         // 设置属性
        newItem.opacity = opacity;
        newItem.blendingMode = blendMode;
        newItem.name = name;
 
        // 删除旧图
        oldItem.remove();
        parentGroup.clipped = true;
 
    }

 
}

// replaceImageByName(); function testing(){     // replaceImageByName('{~
// img}','i-c.png');     replaceImageCover(             getItemsByName('{~
// img}')             ,         'i-c.png'         ); } testing();