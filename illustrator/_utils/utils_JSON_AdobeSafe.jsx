/**
 * JSON_AdobeSafe.jsx
 * ExtendScript 兼容 JSON.stringify（带 Adobe 对象安全支持）
 */
var JSON = JSON || {};

(function () {
    // 检查是否为数组
    function isArray(obj) {
        if (!obj) return false;
        try {
            return typeof obj.length === "number" && !(obj.propertyIsEnumerable("length"));
        } catch (e) {
            return false;
        }
    }

    // 判断是否为 Adobe 对象
    function isAdobeObject(obj) {
        if (!obj || typeof obj !== "object") return false;
        try {
            if (obj.typename) return true;
            if (obj.constructor && obj.constructor.name && obj.constructor.name.indexOf("Adobe") >= 0) return true;
        } catch (e) {}
        return false;
    }

    // 提取安全属性
    function extractAdobeProps(obj) {
        var result = {};
        var safeProps = [];

        try {
            if (app && app.name) {
                var appName = app.name.toLowerCase();
                if (appName.indexOf("illustrator") >= 0) {
                    if (obj.typename === "Document") {
                        safeProps = ["typename", "name", "path", "fullName", "activeLayer"];
                    } else if (obj.typename && obj.typename.match(/Item$/)) {
                        safeProps = ["typename", "name", "visible", "locked", "width", "height", "position", "opacity"];
                    } else {
                        safeProps = ["typename", "name"];
                    }
                } else if (appName.indexOf("photoshop") >= 0) {
                    if (obj.typename === "Document") {
                        safeProps = ["typename", "name", "path", "width", "height", "mode"];
                    } else if (obj.typename === "ArtLayer") {
                        safeProps = ["typename", "name", "visible", "opacity", "kind"];
                    } else if (obj.typename === "LayerSet") {
                        safeProps = ["typename", "name", "visible", "layers"];
                    } else {
                        safeProps = ["typename", "name"];
                    }
                }
            }
        } catch (e) {}

        for (var i = 0; i < safeProps.length; i++) {
            var key = safeProps[i];
            try {
                var val = obj[key];
                if (val !== undefined && typeof val !== "function") {
                    result[key] = val;
                }
            } catch (e) {
                // 跳过无法访问的属性
            }
        }

        return result;
    }

    // 防循环引用
    var seen = [];

    function stringifySafe(obj) {
        if (obj === null) return "null";
        if(obj.typename){
            return '"' + obj.typename + '"';
        }

        var type = typeof obj;

        if (type === "number" || type === "boolean") return String(obj);
        if (type === "string") return '"' + obj.replace(/"/g, '\\"') + '"';
        if (type === "undefined" || type === "function") return undefined;

        if (type === "object") {
            if (seen.indexOf(obj) >= 0) return '"[Circular]"';
            seen.push(obj);

            if (isArray(obj)) {
                var arr = [];
                for (var i = 0; i < obj.length; i++) {
                    arr.push(stringifySafe(obj[i]) || "null");
                }
                return "[" + arr.join(",") + "]";
            }

            // Adobe 对象处理
            if (isAdobeObject(obj)) {
                obj = extractAdobeProps(obj);
            }

            var props = [];
            for (var key in obj) {
                try {
                    var val = stringifySafe(obj[key]);
                    if (val !== undefined) props.push('"' + key + '":' + val);
                } catch (e) {}
            }
            seen.pop();
            return "{" + props.join(",") + "}";
        }

        return '"[Unknown]"';
    }


    // === JSON 序列化 ===
    function toJSONString(obj, indent) {
        if (typeof indent === "undefined") indent = " ";
        var nextIndent = indent + "  ";

        if (obj === null) return "null";
        if (typeof obj === "string") return '"' + obj.replace(/"/g, '\\"') + '"';
        if (typeof obj === "number" || typeof obj === "boolean") return String(obj);

        if (obj.constructor === Array) {
            var arrStr = [];
            for (var i = 0; i < obj.length; i++) {
                arrStr.push(toJSONString(obj[i], nextIndent));
            }
            return "[\n" + nextIndent + arrStr.join(",\n" + nextIndent) + "\n" + indent + "]";
        }

        var keys = [];
        for (var k in obj) {
            keys.push('"' + k + '": ' + toJSONString(obj[k], nextIndent));
        }
        return "{\n" + nextIndent + keys.join(",\n" + nextIndent) + "\n" + indent + "}";
    }



    JSON.stringify = stringifySafe;
    
})();
