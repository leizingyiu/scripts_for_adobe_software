
function arrayLikeToArray(like) {
    return Array.prototype.slice.call(like || []);
}

/** 定义属性
 * @param {Object} obj - 要定义属性的对象
 * @param {String} prop - 要定义的属性名
 * @param {Object} descriptor - 属性描述符
 * @returns {Object} obj - 添加属性后的对象
 * @example
 * var obj = {}
 * ____defineProperty(obj, "prop", {})
 * return obj
 * */
function ____defineProperty(obj, prop, descriptor) {
    if (obj == null) throw new TypeError('Object.defineProperty called on null or undefined');
    if ('value' in descriptor) obj[prop] = descriptor.value;
    // ES3 不能真的控制 writable、configurable，只能用简单标记
    // enumerable 可以通过自定义循环函数来模拟，但 for..in 会仍然枚举
    return obj;
};

// ES3 模拟 Object.defineProperty
if (typeof Object.defineProperty !== 'function') {
    Object.defineProperty = ____defineProperty;
}

if (!Object.defineProperty) {
    alert("当前环境不支持 Object.defineProperty");
    throw "当前环境不支持 Object.defineProperty";
}




/** 在数组中查找指定元素
 * @param {*} element - 要查找的元素
 * @returns {Number} 如果找到元素，返回其在数组中的索引位置；否则返回-1
 */
function ____indexOf(element) {
    for (var i = 0; i < this.length; i++) {
        if (this[i] === element) {
            return i;
        }
    }
    return -1;
}

/** 在 Array.prototype 上添加 indexOf 方法
 * 在数组中查找指定元素的索引位置
 * @param {*} element - 要查找的元素
 * @returns {Number} 如果找到元素，返回其在数组中的索引位置；否则返回-1
 */
if (!Array.prototype.indexOf && Object.defineProperty) {
    Object.defineProperty(Array.prototype, "indexOf", {
        value: ____indexOf,
        writable: true,
        configurable: true,
        enumerable: false
    })
}



/** 在数组中查找满足条件的元素
 * @param {Function} callback - 回调函数，返回 true 表示该元素满足条件
 * @param {*} thisArg - 回调函数的 this 对象
 * @returns {Array} 满足条件的元素组成的数组
 */
function ____map(callback, thisArg) {
    var T, A, k;

    if (this == null) {
        throw new TypeError('this is null or not defined');
    }

    var O = Object(this);
    var len = O.length >>> 0;

    if (typeof callback !== 'function') {
        throw new TypeError(callback + ' is not a function');
    }

    if (arguments.length > 1) {
        T = thisArg;
    }

    A = new Array(len);
    k = 0;

    while (k < len) {
        var kValue, mappedValue;

        if (k in O) {
            kValue = O[k];
            mappedValue = callback.call(T, kValue, k, O);
            A[k] = mappedValue;
        }
        k++;
    }

    return A;
}
// 在 Array.prototype 上添加 map 方法
if (!Array.prototype.map && Object.defineProperty) {
    Object.defineProperty(Array.prototype, "map", {
        value: ____map,
        writable: true,
        configurable: true,
        enumerable: false
    });
};



/** 在数组中查找满足条件的元素
* @param {Function} callback - 回调函数，返回 true 表示该元素满足条件
* @param {*} thisArg - 回调函数的 this 对象
* @returns {Array} 满足条件的元素组成的数组
*/
function ____filter(callback, thisArg) {
    var T, A, k;

    if (this == null) {
        throw new TypeError('this is null or not defined');
    }

    var O = Object(this);
    var len = O.length >>> 0;

    if (typeof callback !== 'function') {
        throw new TypeError(callback + ' is not a function');
    }

    if (arguments.length > 1) {
        T = thisArg;
    }

    A = [];
    k = 0;

    while (k < len) {
        var kValue;

        if (k in O) {
            kValue = O[k];
            // 调用回调函数，如果返回 true 则保留该元素
            if (callback.call(T, kValue, k, O)) {
                A.push(kValue);
            }
        }
        k++;
    }

    return A;
}
// 在 Array.prototype 上添加 filter 方法
if (!Array.prototype.filter && Object.defineProperty) {
    Object.defineProperty(Array.prototype, "filter", {
        value: ____filter,
        writable: true,
        configurable: true,
        enumerable: false
    });
}



/** 数组元素求和
 * @param {Array} array - 要处理的数组
 * @param {*} initialValue - 初始值
 * @returns {*} 处理结果
 */
function ____reduce(callback /*, initialValue*/) {
    if (this == null) {
        throw new TypeError('Array.prototype.reduce called on null or undefined');
    }
    if (typeof callback !== 'function') {
        throw new TypeError(callback + ' is not a function');
    }

    var O = this; // ES3 不支持 Object(this) 很多情况下也没问题
    var len = O.length >>> 0; // 转成正整数
    var k = 0;
    var accumulator;

    // 判断是否有初始值
    if (arguments.length > 1) {
        accumulator = arguments[1];
    } else {
        // 找第一个非空元素作为初始值
        var found = false;
        while (k < len && !found) {
            if (k in O) {
                accumulator = O[k];
                found = true;
            }
            k++;
        }
        if (!found) {
            throw new TypeError('Reduce of empty array with no initial value');
        }
    }

    // 执行 reduce
    for (; k < len; k++) {
        if (k in O) {
            accumulator = callback(accumulator, O[k], k, O);
        }
    }

    return accumulator;
}

// 在 Array.prototype 上添加 reduce 方法
if (!Array.prototype.reduce && Object.defineProperty) {
    Object.defineProperty(Array.prototype, "reduce", {
        value: ____reduce,
        writable: true,
        configurable: true,
        enumerable: false
    });
}


/** 数组去重
 * @returns {Array} 去重后的数组
 */
function ____unique() {
    var result = [];
    for (var i = 0; i < this.length; i++) {
        var exists = false;
        for (var j = 0; j < result.length; j++) {
            if (result[j] === this[i]) { // 基本类型比较
                exists = true;
                break;
            }
        }
        if (!exists) {
            result.push(this[i]);
        }
    }
    return result;
}
// 在 Array.prototype 上添加 unique 方法
if (!Array.prototype.unique && Object.defineProperty) {
    Object.defineProperty(Array.prototype, 'unique', {
        value: ____unique,
        writable: true,
        configurable: true,
        enumerable: false // 不让它出现在 for...in 循环里
    });
}












// function test(){// 测试
//     var arr = [1,2,2,3,3,4,1];
//     var uniqueArr = arr.unique();
//     var reduceArr = arr.reduce(function(pre, cur) {return pre+cur},0);
//     var filterArr = arr.filter(function(item, index, arr) {        return item<=2;    });
//     var mapArr = arr.map(function(item, index, arr) {        return item+1;    });
//     var idx = arr.indexOf(2);

//     alert(
//         [
//             ['unique',uniqueArr].join(':'),
//             ['reduce',reduceArr].join(':'),
//             ['filter',filterArr].join(':'),
//             ['map',mapArr].join(':'),
//             ['indexOf',idx].join(':')
//         ].join('\n')
//     );
// }

// test();