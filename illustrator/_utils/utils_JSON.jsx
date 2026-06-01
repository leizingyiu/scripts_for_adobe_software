var JSON=function(){};
var errStep=0;
(function ensureJSON() {
  
function isArray(obj) {
    if (obj == null || typeof obj !== "object") {
        return false;
    }
    
    // 只检测最基本的特征
    if (typeof obj.length === "number" && obj.length >= 0) {
        // 如果是空对象但有 length 属性，很可能是数组
        var count = 0;
        for (var key in obj) {
            if (obj.hasOwnProperty(key)) {
                count++;
            }
        }
        // 如果属性数量大致等于 length + 1（包括 length 本身），可能是数组
        if (count === obj.length + 1 || count === obj.length) {
            return true;
        }
    }
    
    return false;
}

  function jsonStringify(obj) {
    if (obj === null) {
        return "null";
    }
    var type = typeof obj;

    if (type === "number" || type === "boolean") {
        return String(obj);
    }

    if (type === "string") {
        return '"' + obj.replace(/"/g, '\\"') + '"';
    }

    if (type === "object") {
        // 数组
        if (isArray(obj)) {
            var arr = [];
            for (var i = 0; i < obj.length; i++) {
                arr.push(jsonStringify(obj[i]));
            }
            return "[" + arr.join(",") + "]";
        }
        // 普通对象
        var props = [];
        for (var key in obj) {
            if (obj.hasOwnProperty(key)) {
                props.push('"' + key + '":' + jsonStringify(obj[key]));
            }
        }
        return "{" + props.join(",") + "}";
    }

    return undefined; // 不支持的类型
}

// 完全避免 hasOwnProperty 的版本
function jsonStringifySafe(obj) {
    if (obj === null) {
        return "null";
    }
    
    var type = typeof obj;

    if (type === "number" || type === "boolean") {
        return String(obj);
    }

    if (type === "string") {
        return '"' + obj.replace(/"/g, '\\"') + '"';
    }

    if (type === "object") {
        // 尝试作为数组处理
        try {
            if (typeof obj.length === "number" && obj.length >= 0) {
                var arr = [];
                var successCount = 0;
                
                // 尝试访问前几个元素来判断是否是数组
                for (var i = 0; i < obj.length; i++) {
                    try {
                        arr.push(jsonStringifySafe(obj[i]));
                        successCount++;
                    } catch (e) {
                        arr.push("null");
                    }
                }
                
                // 如果成功访问了大部分元素，认为是数组
                if (successCount > 0 || obj.length === 0) {
                    return "[" + arr.join(",") + "]";
                }
            }
        } catch (e) {
            // 继续作为普通对象处理
        }
        
        // 作为普通对象处理
        var props = [];
        try {
            for (var key in obj) {
                try {
                    var value = obj[key];
                    var valueType = typeof value;
                    
                    if (valueType !== "function" && valueType !== "undefined") {
                        props.push('"' + key + '":' + jsonStringifySafe(value));
                    }
                } catch (e) {
                    // 跳过无法访问的属性
                    continue;
                }
            }
        } catch (e) {
            return "{}";
        }
        return "{" + props.join(",") + "}";
    }

    return undefined;
}

JSON.stringify=jsonStringifySafe;

})();