function isEmptyContent(str) {
    if (typeof str !== "string") 
        str = String(str);
    
    str = str.replace(/^\s+|\s+$/g, '');

    if (str === "" || str === "null" || str === "undefined" || str === "NaN" || str === "FALSE" || str === "false" || (str.toLowerCase && str.toLowerCase() === "false")) {
        return true;
    }
    return false;
}

/* warnning message : alert() */
function warn(msg) {
    if (!confirm('warn : ' + msg + '\n是否继续？')) {
        throw new Error('用户已取消 \n@warn : ' + msg);
    };
};


/** 显示确认对话框，询问用户是否继续执行
 * 如果用户选择取消，则抛出错误
 * @param {string} msg - 要显示给用户的消息内容
 * @returns {boolean} 当用户确认时返回true，取消时抛出错误
 * @throws {Error} 当用户选择取消时抛出'用户已取消'错误
 */
function carryOn(msg) {
    var result = confirm(msg + '\n继续吗？');
    if (!result) {
        throw new Error('用户已取消');
        return false;
    }
    return result;
}


function showArguments(_arguments, showFn) {
    var showFn = showFn || alert;
    var result = "";
    for (var i = 0; i < _arguments.length; i++) {
        result += "Argument " + i + ": " + _arguments[i] + "\n";
    }
    showFn('showArguments: \n' + result);
}

function scriptFolder($_fileName) {
    return (new File($_fileName)).parent;
}

function stepAppRedraw() {
    var enableStepRedraw = enableStepRedraw || false;
    if (enableStepRedraw === true) {
        app.redraw();
    }
}

function log(e, msg, folder, filename, T, maxSize, skipLines) {
    var enableLog = enableLog || false;
    if (!enableLog) return;

    var enableCarryOnWhenLog = enableCarryOnWhenLog || false;
    if (enableCarryOnWhenLog === true && e) {
        carryOn('遇到错误' + e + (msg ? '\n提示为' + msg : ''));
    }

    var enableReverseLog = enableReverseLog || false;

    var now = new Date();
    var timestamp =
        now.getFullYear() + "-" +
        ("0" + (now.getMonth() + 1)).slice(-2) + "-" +
        ("0" + now.getDate()).slice(-2) + " " +
        ("0" + now.getHours()).slice(-2) + ":" +
        ("0" + now.getMinutes()).slice(-2) + ":" +
        ("0" + now.getSeconds()).slice(-2);

    var currentScript = new File($.fileName);
    var scriptFolder = currentScript.parent;

    folder = folder || scriptFolder;
    filename = filename || "ai_error_log.txt";
    msg = msg || '';

    var logFile = new File(folder + "/" + filename);

    if (enableReverseLog) {
        maxSize = (maxSize || 500) * 1024;
        skipLines = (typeof skipLines === "number") ? skipLines : 2;

        var newEntry = timestamp + ": " + e + (e && e.line ? " at line " + e.line : "") + "\n" + msg + "\n\n";

        var oldContent = "";
        if (logFile.exists) {
            logFile.open("r");
            logFile.encoding = "UTF-8"; // ✅ 设置读取编码
            if (logFile.length > maxSize) {
                var allLines = [];
                while (!logFile.eof) {
                    allLines.push(logFile.readln());
                }
                logFile.close();

                var filtered = [];
                var timeRegex = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}:/;
                var T_time = T ? new Date(T).getTime() : null;

                for (var i = 0; i < allLines.length; i += (skipLines + 1)) {
                    var line = allLines[i];
                    if (!timeRegex.test(line)) continue;

                    if (T_time) {
                        var lineTime = new Date(line.substring(0, 19)).getTime();
                        if (!isNaN(lineTime) && lineTime < T_time) {
                            break;
                        }
                    }
                    filtered.push(line);
                }

                oldContent = (filtered.length > 0)
                    ? filtered.join("\n") + "\n"
                    : "log文件太大，请手动清理。\n";

            } else {
                oldContent = logFile.read();
                logFile.close();
            }
        }

        logFile.open("w");
        logFile.encoding = "UTF-8"; // ✅ 设置写入编码
        logFile.write(newEntry + oldContent);
        logFile.close();

    } else {
        logFile.open("a");
        logFile.encoding = "UTF-8"; // ✅ 设置写入编码
        var lineInfo = (e && e.line) ? " at line " + e.line : "";
        logFile.writeln(timestamp + ": " + e + lineInfo + "\n" + msg + "\n");
        logFile.close();
    }
}



function revealFolder(){     
    if(typeof app.activeDocument =='undefined'){
        throw new Error("revealFolder() : 当前未打开文档");
    }
    app.activeDocument.path;
    app.activeDocument.path.execute();
}