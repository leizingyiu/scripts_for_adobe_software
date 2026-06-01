#include "_utils/prototype_Array.jsx"
#include "_utils/utils.jsx"

#include "_utils/utils_JSON_AdobeSafe.jsx"

#include "_utils/UI_saveOrContinue.jsx" 
#include "_utils/UI_checkAskAndUnlockUnhide.jsx"
#include "_utils/UI_func_uiDialog.jsx"

#include "_utils/func_layout_alignChildren.jsx"
#include "_utils/func_layout_alignToTarget.jsx"
#include "_utils/func_layout_distributeChildren.jsx"
#include "_utils/func_layout_spaceChildren.jsx"
#include "_utils/func_fillText.jsx"
#include "_utils/func_imgReplace.jsx"

#include "_utils/utils_artboard.jsx"  
#include "_utils/utils_group.jsx"  

#include "_utils/utils_select.jsx"

#include "_utils/excelContentFiller - func_queryClosure.jsx"
#include "_utils/excelContentFiller - UI_tabTextToObjArray.jsx"

var global_config = {
    artboard_gap: 400,                  // 画布间距  
    artboard_copy_direction: "right",   // left,right,top/up,bottom/down
    log:false,                          // 记录到日志中；
    reverseLog:false,                   // 反向记录，最新内容在最上面，可能导致缓慢；
    carryOnWhenLog:false,               // 记录时，弹出 carryOn 提示框
    redraw:true,                        // 是否在步骤中重绘
}



//激活配置项
var enableLog=global_config.log || false ;
var enableStepRedraw = global_config.redraw || false ;
var enableReverseLog = global_config.reverseLog || false ;
var enableCarryOnWhenLog = global_config.carryOnWhenLog || false ;


// excelContentFiller 主功能函数
function main_fn(datas) {

    log('excelContentFiller ： main_fn start');

    var testingData = {
        "文本 1": '文1文1',
        "文本 2": '本2本本2',
        "文本 3": '本3本本3\n本3本3本',
        show1: false,
        show2: true,
        img: "i-c",
        source_artboard_name: '模板 1'
    };
    var testingData2 = {
        "文本 1": '文1文1--2',
        "文本 2": '本2本本2--2',
        "文本 3": '本3本本3\n本3本3本',
        show1: false,
        show2: true,
        img: "i-y",
        source_artboard_name: '模板 1'
    };

    datas = datas || [testingData, testingData2];

    var artboardAccaptKeys = ['source_artboard_name', '画板', '模板画板', '模板'];

    var namesDict = {
        tempFill: '模板填充',
        fnFill: '函数填充',
        txtFill: '文本填充'
    }

    var regDict = {
        // 文本填充:/((_ .+)|(`.*\{_ [^\}]*\}.*`)|(_\([a-zA-Z_$][a-zA-Z0-9_$]*\) .+))/, //
        // 文本填充的所有reg
        文本填充: /((_ .+)|(`.*\{_ [^\}]*\}.*`)|(_\([^\)]*\) .+))/, // 文本填充的所有reg
        模板填充: /`.*\{_ [^\}]*\}.*`/, // 使用对象名中的文本模板填充文本
        // 函数填充: /_\([a-zA-Z_$][a-zA-Z0-9_$]*\) .+/,    // 使用特定函数填充文本
        函数填充: /_\([^\)]*\) .+/, // 使用特定函数填充文本
        切换可见: /\+ .+/, // 切换图层是否可见
        图片填充: /~ .+/, // 填充文档所在文件夹中指定名称的图片
        对齐: /[\^v<>]=[\^v<>]/, // item 都向指定方向对齐
        排列: /[\^v<>]{2}/, // (左item左 到 右item左)
        间距: /[\^v<>]/, //(左item右 到 右item左)
    };
    var argDict = { // 识别到的多种形式，整理成指定形式的参数
        文本填充: ['函数填充', '模板填充']
    };

    var argArr_ele_of_fn_dict = {
        模板填充: '文本填充',
        函数填充: '文本填充'
    };

    var filterWithoutDict = {/* keep:['remove1','remove2']*/
        文本填充: [
            '函数填充', '模板填充'
        ],
        对齐: ['间距'],
        排列: ['间距']
    };

    var fnDict = {
        文本填充: function (item, data, argArr) {
            log('文本填充() start : ' + item.name);
            argArr = argArr || [];
            var fnreg = /.*\{_\(([^\)]*)\) .+/g;
            var _fnDict = {

                "固定行数填充": fillAndKeepLinesForTxtAvatar,
                "行": fillAndKeepLinesForTxtAvatar,
                "fillAndKeepLinesForTxtAvatar": fillAndKeepLinesForTxtAvatar,
                "forceLines": fillAndKeepLinesForTxtAvatar,

                "文本框填充": fillTxtInFrame,
                "框": fillTxtInFrame,
                "fillTxtInFrame": fillTxtInFrame,
                "inFrame": fillTxtInFrame,

                "单行文本框只变宽":fillAreaTxtWider,
                "单行框变宽":fillAreaTxtWider,
                "fillAreaTxtWider":fillAreaTxtWider,
                "fillTxtWider":fillAreaTxtWider,


                "文本填充": fillText,
                "文": fillText,
                "fillText": fillText,

            };

            var content = '';
            var _isTemplate = false;
            var _isSpecificFunction = false;

            for (var i = 0; i < argArr.length; i++) {
                _isTemplate = argArr[i] == '模板填充'
                    ? true
                    : _isTemplate;
                _isSpecificFunction = argArr[i] == '函数填充'
                    ? true
                    : _isSpecificFunction;
            }
            log("文本填充()：模板：" + _isTemplate + " ；特定函数：" + _isSpecificFunction+' // '+item.name);

            if (_isTemplate) {
                content = item
                    .name
                    .replace(/.*`([^`]*)`.*$/, '$1')
                    .replace(/\{_ ([^\}]*)\}/g, function (match, p1) {
                        return data[p1]
                            ? data[p1]
                            : '!!!fnDict - 文本填充 - isTemplate - 匹配失败!!!';
                    });
            } else {
                content = item
                    .name
                    .replace(/.*\{_\S* ([^\}]*)\}.*$/, '$1')
                    .replace(/^(.*)$/, function (match, p1) {
                        return data[p1]
                            ? data[p1]
                            : '!!!fnDict - 文本填充 - default - 匹配失败!!!\n\n' + [item.name, data, argArr].join('\n');
                    });
            }

            log("文本填充 end get content ： content： " + content +' // '+item.name );

            var fn = fillText,
                fnName = 'fillText';
            if (_isSpecificFunction) {
                var _fnName = item
                    .name
                    .replace(fnreg, '$1');
                log('文本填充 get function ： ' + _fnName);

                if (_fnDict[_fnName]) {
                    fnName = _fnName;
                    fn = _fnDict[fnName];
                    
                    log('文本填充 get function ： ' + fnName  );

                } else {
                    warn('!!!fnDict - 文本填充 - isFunction - 函数<<' + _fnName + '>>不存在!!!\n\n' + [item.name, data, argArr, _isTemplate, _isSpecificFunction].join('\n'));
                }
            }
            log("文本填充 end get fn ： "+item.name + '\nMatched content: ' + content + '\n fn:' + fnName + '\n arg:' + argArr);

            try {
                fn(item, content);
            } catch (e) {
                throw new Error('!!!文本填充 - 执行函数 ' + fnName + ' 失败!!!\n' + e + '\n _main_getAndRun.js');
            }
            log('文本填充: end'+ item.name );
        },
        切换可见: function (item, data, argArr) {
            log('切换可见: start' + item.name);
            var valueKey = item
                    .name
                    .replace(/.*\{\+ ([^\}]*)\}.*$/, '$1'),
                value = data[valueKey];
            if (value == undefined) {
                warn('!!!切换可见 - 未找到对应的值!!!');
                return;
            }
            value = value.toLowerCase() == 'true'
                ? true
                : (value.toLowerCase() == 'false'
                    ? false
                    : value);
            if (item.typename == "Layer") {
                item.visible = Boolean(value);
            } else {
                item.hidden = !Boolean(value);
            }
            log('切换可见: end' + item.name);
        },

        图片填充: function (item, data, argArr) {

            log('图片填充: start' + item.name);

            var valueKey = item
                    .name
                    .replace(/.*\{~ ([^\}]*)\}.*$/, '$1'),
                value = data[valueKey];

            log('图片填充: ' + item.name + ' - ' + valueKey + " - " + value);

            replaceImageCover(item, value);
            log('图片填充: end' + item.name);

        },
        对齐: function (item, data, argArr) {
            log("对齐() start:" + item.name);

            var targetItems = findChildrenByName(item, /\{=\}/);
            var targetItem = targetItems.length > 0
                ? targetItems[0]
                : false;
            var _alignment = item
                .name
                .match(/[<^>v]=[<^>v]/);

            if (!_alignment) {
                warn('!!!对齐 - 未识别到有效的对齐方式!!!\n\n' + [item.name, data, argArr, _alignment].join('\n'));
                return;
            }
            _alignment = _alignment[0];
            // carryOn(_alignment);

            var alignment = '';
            switch (_alignment) {
                case '<=<':
                    alignment = 'left';
                    break;
                case '>=>':
                    alignment = 'right';
                    break;
                case 'v=v':
                    alignment = 'bottom';
                    break;
                case '^=^':
                    alignment = 'top';
                    break;
                default:
                    warn('!!!对齐 - 未识别到有效的对齐方式!!!\n\n' + [item.name, data, argArr, _alignment].join('\n'));
                    return;
            }
            log("对齐() ： hasTargetItem?:" + Boolean(targetItem) + '; \n 对齐方式： ' + alignment + "\n " + item.name);

            if (targetItem != false) {
                alignChildrenToTarget(item, targetItem, alignment);
            } else {
                alignChildren(item, alignment);
            }

            log("对齐() end" + item.name);

        },
        排列: function (item, data, argArr) {
            log("排列() start : " + item.name);

            var _direction = item
                .name
                .match(/[<^>v]{2}/);
            if (!_direction) {
                warn('!!!排列 - 未识别到有效的排列方式!!!\n\n' + [item.name, data, argArr, _direction].join('\n'));
                return;
            }
            _direction = _direction[0];
            var direction = '';
            switch (_direction) {
                case '<<':
                    direction = 'left';
                    break;
                case '>>':
                    direction = 'right';
                    break;
                case 'vv':
                    direction = 'bottom';
                    break;
                case '^^':
                    direction = 'top';
                    break;
                default:
                    warn('!!!排列 - 未识别到有效的排列方式!!!\n\n' + [item.name, data, argArr, _direction].join('\n'));
                    return;
            }

            log('排列() \n方向：' + direction + '\n组内元素数：' + item.pageItems.length + '\n组名：' + item.name);

            distributeChildren(item, direction);
            log("排列() end : " + item.name);
        },
        间距: function (item, data, argArr) {
            log("间距() start : " + item.name);

            var _direction = item
                .name
                .match(/[<^>v]/);
            if (!_direction) {
                warn('!!!间距 - 未识别到有效的排列方式!!!\n\n' + [item.name, data, argArr, _direction].join('\n'));
                return;
            }
            _direction = _direction[0];
            var direction = '';
            switch (_direction) {
                case '<':
                    direction = 'left';
                    break;
                case '>':
                    direction = 'right';
                    break;
                case 'v':
                    direction = 'bottom';
                    break;
                case '^':
                    direction = 'top';
                    break;
                default:
                    warn('!!!间距 - 未识别到有效的排列方式!!!\n\n' + [item.name, data, argArr, _direction].join('\n'));
                    return;
            }

            log('间距() \n方向：' + direction + '\n组内元素数：' + item.pageItems.length + '\n组名：' + item.name);
            spaceChildren(item, direction);

            log("间距() end : " + item.name);

        }
    }
    var fnsOrder = [
        '切换可见',
        '模板填充',
        '文本填充',
        '图片填充',
        '对齐',
        '排列',
        '间距'
    ];

    var artboardAccaptKeys = artboardAccaptKeys || ['source_artboard_name', '画板', '模板画板'];

    var doc = app.activeDocument;
    if (!doc) {
        warn("请先打开一个 Illustrator 文档。");
    }

    var ABs,
        AB,
        nAB,
        AB_count = {};

    log('excelContentFiller main_fn 正在处理文档： '+doc.name);

    // 根据数据复制画板和内容
    for (var i = 0; i < datas.length; i++) {

        var source_artboard_name;

        for (var j = 0; j < artboardAccaptKeys.length; j++) {
            if (typeof datas[j][artboardAccaptKeys[j]] !== 'undefined') {
                source_artboard_name = datas[j][artboardAccaptKeys[j]];
                break;
            }
        }

        if (!source_artboard_name) {
            throw new Error('来源数据中没有指定画板 @ _main_getAndRun.js');
        }

        var useFirstAB = false;
        if (typeof source_artboard_name === 'undefined') {
            useFirstAB = carryOn('来源数据中没有指定画板，是否使用 第一个画板？');
        }

        log('即将复制画板：' + source_artboard_name);

        if (useFirstAB) {
            AB = doc.artboards[0];
            log('使用第一个画板：' + AB.name + ". _ AB = doc.artboards[0]; ");
        } else {
            ABs = getArtboardsByName(doc, source_artboard_name);

            if (!ABs.length) {
                if (carryOn('未找到 artboard：' + source_artboard_name)) {
                    continue;
                } else {
                    throw new Error('未找到 artboard：' + source_artboard_name + ' @ getAndRun.js');
                }
            }
            if (ABs.length > 1) {
                if (carryOn('找到多个 artboard：' + source_artboard_name)) {
                    continue;
                } else {
                    throw new Error('找到多个 artboard：' + source_artboard_name + ' @ getAndRun.js');
                }
            }

            AB = ABs[0];
            log('使用 artboard：' + AB.name + ". _  AB = ABs[0]; ");
        }

        if (typeof AB_count[source_artboard_name] == 'undefined') {
            AB_count[source_artboard_name] = 0;
        }

        if (AB_count[source_artboard_name] != 0) {
            log('复制画板 - start ：' + AB.name);

            nAB = duplicate_AB_and_CNT(AB, 1, (typeof global_config != 'undefined' && typeof global_config.artboard_gap != 'undefined'
                ? global_config.artboard_gap
                : 400), (typeof global_config != 'undefined' && typeof global_config.artboard_copy_direction != 'undefined'
                ? global_config.artboard_copy_direction
                : "right"));
            log('复制画板 - end ：' + AB.name + ':' + indexOfArtboard(AB) + ' -> ' + nAB.name + ':' + indexOfArtboard(nAB));
        } else {
            nAB = AB;
            log('直接使用原画板：' + AB.name);
        }

        datas[i].artboard = nAB;
        log(" datas[" + i + "].artboard = nAB; - " + nAB.name + "\n AB_count:"+AB_count[source_artboard_name] + ' ; 画板名：' + source_artboard_name);

        AB_count[source_artboard_name] = (AB_count[source_artboard_name] || 0) + 1;      

        log('in artboart copying end  : ' + indexOfArtboard(datas[i].artboard) + ' : ' + datas[i].artboard.name);

      
    }

    stepAppRedraw();

    var unlockedItems = [],
        noneFnItems = [],
        processedItems = [],
        processeGroups = [];

    // 开始按照 datas 进行处理：
    for (var k = 0; k < datas.length; k++) {

        var data = datas[k];

        AB = data.artboard;

        log(k + ': before  process :' + AB.name);

        processArtboardsByLayers(AB, function (_items) {

            log(k + ': start process :' + AB.name);

            var AB_group = group(_items);

            AB_group.name = '__组__AB' + Math.random();

            var itemGroups = queryGet(AB_group, false, collectPageItemsByRoots);
            // getPageItemsFrom_qwen / itemGroups.keys = sortByPattern(itemGroups.keys,
            // regDict, fnsOrder);

            var itemGroupsKeys = sortByPattern(itemGroups.keys, regDict, fnsOrder);

            // alert([fnsOrder,itemGroups.keys, itemGroupsKeys].map(function(___){return
            // ___.join(',')}).join('\n'));

            for (var j = 0; j < itemGroupsKeys.length; j++) {

                var matchKey = itemGroupsKeys[j];
                var items = itemGroups[matchKey];

                log(j + ': in process ' + matchKey + '; ' + items.length + ' items');

                // alert([matchKey,
                // itemGroups[matchKey].length,indexOfArtboard(data.artboard)].join('\n'));

                for (var i = 0; i < items.length; i++) {
                    var item = items[i];

                    log(i + ': in item process ' + matchKey + '; ' + item.name);

                    // alert([i, j, k, matchKey, item.name].join('\n'));

                    var _fn = arrayRegDictMatch([matchKey], regDict);
                    var __fn = _fn;
                    var argArr = [];

                    if (_fn.length == 0) {
                        noneFnItems.push(item);
                        continue;
                    } else {
                        _fn = _fn[0];
                    }

                    for (_k in argArr_ele_of_fn_dict) { // 识别出操作函数以及参数
                        if (_fn == argArr_ele_of_fn_dict[_k]) {
                            _fn = argArr_ele_of_fn_dict[_k];
                        }
                        if (__fn.indexOf(_k) != -1) {
                            argArr.push(_k);
                        }
                    }

                    log('before fn() : fn and argArr : ' + [_fn, argArr].join(' --- '));

                    // stepAppRedraw(); alert([k, j, i, _fn, argArr, item.name, data].join('\n'));
                    // alert([     "indexOfArtboard(data.artboard) :
                    // ",indexOfArtboard(data.artboard),         "\n",     "_fn : ",_fn, "\n",
                    // "item.name : ",item.name,           "\n",     "matchKey : ",matchKey, "\n",
                    //   "data[matchKey] : ",data[matchKey], "\n",     "argArr : ",argArr,
                    // "\n", ].join('\n'));

                    fnDict[_fn](item, data, argArr);

                    log('after fn() : fn and argArr : ' + [_fn, argArr].join(' --- '));

                    // carryOn(['6',     _fn,     item.name,     matchKey,     toJSONString(data),
                    // argArr, ].join('\n'));

                    processedItems.push(item);

                    log(i + ': in item process done ' + matchKey + '; ' + item.name);

                }
                stepAppRedraw();
            }

            processeGroups.push(AB_group);

            AB_group.locked = true;

            log(k + ': finish process ' + AB.name);

        })

        log(k + ': after process :' + AB.name);
        
        stepAppRedraw();
    }

    
    log('process done , start unlog them');

    for (i = 0; i < processeGroups.length; i++) {
        var _group = processeGroups[i];
        _group.locked = false;
        ungroup(_group);
    }

    log('unlog done');
    log('final log start');

    processedItems = processedItems.unique();
    unlockedItems = unlockedItems.unique();
    noneFnItems = noneFnItems.unique();

    var finalLog = function () {
        uiDialog('alert', ('处理完成，共处理 ' + itemGroups.keys.length + ' 类图层。\n' + itemGroups.keys.join(' | ') + '\n___\n处理了 ' + processedItems.length + ' 个图层。\n' + groupItemsByTopLayer(processedItems).str + '\n___\n锁定的图层有 ' + unlockedItems.length + ' 个。\n' + (unlockedItems.length > 0
            ? ('锁定的图层有：\n' + groupItemsByTopLayer(unlockedItems).str)
            : '') + '\n___\n未识别到任何可用操作的图层有 ' + noneFnItems.length + ' 个。\n' + (noneFnItems.length > 0
            ? ('未识别到任何可用操作的图层有：\n' + groupItemsByTopLayer(noneFnItems).str)
            : '')), {width: 500});
    };

    var falseToShowDetail = confirm('处理完成，共处理 ' + itemGroups.keys.length + ' 类图层。\n' + itemGroups.keys.join(' | ') + '\n___\n处理了 ' + processedItems.length + ' 个图层。\n点击确定退出，点击取消查看详细结果');
    if (!falseToShowDetail) {
        finalLog();
    }
    log(' main_fn done');
}

/** 匹配数组元素与正则表达式字典
 * @param {Array} arr - 要检查的字符串数组
 * @param {Object} dict - 包含正则表达式的对象，键为标识符，值为正则表达式
 * @returns {Array} 包含匹配成功的dict键名的数组
 */
function arrayRegDictMatch(arr, dict) {
    var result = [];
    for (var regKey in dict) {
        for (var j = 0; j < arr.length; j++) {
            var pattern = arr[j];
            if (dict[regKey].test(pattern)) {
                result.push(regKey);
            }
        }
    }
    return result;
}

function sort_A_By_B(a, b) {
    // 创建一个对象来存储 b 数组中元素的顺序索引
    var orderMap = {};
    for (var i = 0; i < b.length; i++) {
        orderMap[b[i]] = i;
    }

    // 过滤出在 b 中存在的元素
    var existingInB = [];
    var notInB = [];

    for (var j = 0; j < a.length; j++) {
        if (orderMap.hasOwnProperty(a[j])) {
            existingInB.push(a[j]);
        } else {
            notInB.push(a[j]);
        }
    }

    // 对在 b 中存在的元素按照 b 的顺序排序
    existingInB
        .sort(function (x, y) {
            return orderMap[x] - orderMap[y];
        });

    // 返回排序后的数组（不在 b 中的元素放在最后）
    return existingInB.concat(notInB);
}

/**
 * 根据指定的模式对数组进行排序
 * @param {Array} A - 需要排序的数组
 * @param {Object} B - 包含名称和正则表达式映射的对象，用于匹配数组元素
 * @param {Array} C - 定义排序顺序的数组，元素为B中的名称
 * @returns {Array} 排序后的新数组
 */
function sortByPattern(A, B, C) {
    // 创建数组副本，避免修改原数组
    var sortedArray = A.slice();

    // 创建一个映射，将元素映射到对应的 name
    var elementToName = {};

    // 建立元素与名称的映射关系
    for (var name in B) {
        if (B.hasOwnProperty(name)) {
            var reg = B[name];
            // 遍历数组 sortedArray，用正则匹配元素
            for (var i = 0; i < sortedArray.length; i++) {
                if (reg.test(sortedArray[i])) {
                    elementToName[sortedArray[i]] = name;
                    // 注意：不 break，因为可能有多个元素匹配同一个模式
                }
            }
        }
    }

    // 创建 C 的顺序映射
    var orderMap = {};
    for (var j = 0; j < C.length; j++) {
        orderMap[C[j]] = j;
    }

    // 按照指定顺序对数组进行排序
    sortedArray
        .sort(function (x, y) {
            var nameX = elementToName[x];
            var nameY = elementToName[y];

            if (nameX && nameY) {
                return orderMap[nameX] - orderMap[nameY];
            } else if (nameX && !nameY) {
                return -1;
            } else if (!nameX && nameY) {
                return 1;
            } else {
                return 0;
            }
        });

    return sortedArray;
}


// 测试函数
function mainTesting(datas) {
    var datas = datas || [
        {
            "文本 1": "aaa",
            "文本 2": "bbb",
            "文本 3": "ccc\nc- c - c",
            "img": "i-y.png",
            "show1": "TRUE",
            "show2": "FALSE",
            "source_artboard_name": "模板 1"
        }, {
            "文本 1": "aaa1",
            "文本 2": "bbb2",
            "文本 3": "ccc3\n    c - c - c",
            "img": "i-m.png",
            "show1": "FALSE",
            "show2": "TRUE",
            "source_artboard_name": "模板 1"
        }, {
            "文本 1": "aaa2",
            "文本 2": "bbb3",
            "文本 3": "ccc4\n    c - c - c",
            "img": "i-k.png",
            "show1": "FALSE",
            "show2": "TRUE",
            "source_artboard_name": "模板 1"
        }, {
            "文本 1": "aaa3",
            "文本 2": "bbb4",
            "文本 3": "ccc5\n    c - c - c",
            "img": "i-c.png",
            "show1": "TRUE",
            "show2": "FALSE",
            "source_artboard_name": "模板 1"
        }
    ];

    var docSaved = UI_saveOrContinue();
    if (docSaved) {
        main_fn(datas);
    } else {
        throw "文档未保存，用户已取消。";
    }

}

function main() {
    var docSaved = UI_saveOrContinue();

    var items = app.activeDocument.pageItems; // 你的 item 列表

    var checker = new LockHiddenChecker();
    var checkResult = checker
        .check(items)
        .ask();

    if (docSaved && checkResult) {
        showExcelParserDialog({'根据数据进行填充': main_fn})
    } else {
        throw "文档未保存，用户已取消。";
    }
}

try {
    main();
} catch (e) {
    alert("错误: " + e + "\n 行号: " + e.line);

    log(e, "@ main() - main_getAndRun.js", scriptFolder($.fileName))
    throw new Error("错误: " + e + "\n 行号: " + e.line);
}