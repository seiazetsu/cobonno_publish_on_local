#target photoshop

// ログファイルを設定
var logFile = new File("~/Desktop/photoshop_script_log.txt");
logFile.open('w');
logFile.close();

// ログ出力用関数
function log(message) {
    logFile.open('a'); // 追記モードでファイルを開く
    logFile.writeln(message);
    logFile.close(); // 書き込み後にファイルを閉じる
}

// CSVファイルを開く
var csvFile = File.openDialog("Select the updated CSV file");
if (!csvFile) {
    alert("No file selected");
    log("No file selected");
    exit();
}
csvFile.open('r');

// CSVデータを読み込む
var csvData = [];
while (!csvFile.eof) {
    var line = csvFile.readln();
    log("Read line: " + line); // デバッグ用ログ
    if (line) { // 空行をスキップ
        // 正規表現を使用してフィールドを分割し、引用符で囲まれたフィールドも考慮
        var fields = line.split(',');
        log("Parsed fields before cleaning: " + fields.join(", ")); // デバッグ用ログ
        for (var j = 0; j < fields.length; j++) {
            fields[j] = fields[j].replace(/^"|"$/g, ''); // 各要素から引用符を削除
        }
        log("Parsed fields after cleaning: " + fields.join(", ")); // デバッグ用ログ
        csvData.push(fields);
    }
}
csvFile.close();
log("CSV data loaded: " + csvData.length + " rows");

// ヘッダーを除去
csvData.shift();

// 各行の画像パスをチェック
for (var i = 0; i < csvData.length; i++) {
    var imagePath = csvData[i][3];
    log("Image path " + (i + 1) + ": " + imagePath); // 画像パスのログ
}

// A4サイズの設定
var a4Width = UnitValue(210, "mm");
var a4Height = UnitValue(297, "mm");
var maxSize = 550; // 画像の長辺を128.5mmに設定

// フォントサイズの設定
var dateFontSize = UnitValue(2, "mm");
var titleFontSize = UnitValue(4, "mm");
var captionFontSize = UnitValue(2, "mm");

// 各要素位置
var imageRefX = 0;
var textRefX = 148;

// 新しいドキュメントを作成
var doc = app.documents.add(a4Width, a4Height, 300);
log("New document created");

// データを各アートボードに流し込む
for (var i = 0; i < csvData.length; i += 2) {
    try {
        log("Processing row: " + (i + 1));
        // 上部モーメントの位置設定
        var refY1 = 0;
        var date1 = "制作日:" + csvData[i][0];
        var title1 = csvData[i][1];
        var description1 = csvData[i][2];
        var imagePath1 = csvData[i][3];
        
        // 下部モーメントの位置設定
        var refY2 = 148.5; // 2行目の基準点
        var date2 = "", title2 = "", description2 = "", imagePath2 = "";
        if (i + 1 < csvData.length) {
            date2 = "制作日:" + csvData[i + 1][0];
            title2 = csvData[i + 1][1];
            description2 = csvData[i + 1][2];
            imagePath2 = csvData[i + 1][3];
        }

        log("Image path 1 before validation: " + imagePath1);
        if (!imagePath1 || imagePath1 === "") {
            log("Image path 1 is undefined or empty");
        }

        log("Image path 2 before validation: " + imagePath2);
        if (!imagePath2 || imagePath2 === "") {
            log("Image path 2 is undefined or empty");
        }

        // 上部モーメントのデータ配置
        log("About to place image 1: " + imagePath1);
        placeImage(imagePath1, imageRefX, refY1 + 2, maxSize);
        log("Placed image 1: " + imagePath1);

        log("About to place date 1: " + date1);
        var dateHeight1 = placeText(date1, textRefX, refY1 + 3, dateFontSize, 50, false);
        log("Placed date 1: " + date1);

        log("About to place title 1: " + title1);
        var titleHeight1 = placeText(title1, textRefX, refY1 + 3 + dateHeight1 + 2, titleFontSize, 50, true);
        log("Placed title 1: " + title1);

        log("About to place description 1: " + description1);
        placeText(description1, textRefX, refY1 + 3 + dateHeight1 + 2 + titleHeight1 + 5, captionFontSize, 50, false);
        log("Placed description 1: " + description1);

        // 下部モーメントのデータ配置
        if (i + 1 < csvData.length) {
            log("About to place image 2: " + imagePath2);
            placeImage(imagePath2, imageRefX, refY2 + 2, maxSize);
            log("Placed image 2: " + imagePath2);

            log("About to place date 2: " + date2);
            var dateHeight2 = placeText(date2, textRefX, refY2 + 3, dateFontSize, 50, false);
            log("Placed date 2: " + date2);

            log("About to place title 2: " + title2);
            var titleHeight2 = placeText(title2, textRefX, refY2 + 3 + dateHeight2 + 2, titleFontSize, 50, true);
            log("Placed title 2: " + title2);

            log("About to place description 2: " + description2);
            placeText(description2, textRefX, refY2 + 3 + dateHeight2 + 2 + titleHeight2 + 5, captionFontSize, 50, false);
            log("Placed description 2: " + description2);
        }

        // ページ保存
        log("About to save page");
        var pageNum = ("000" + Math.ceil((i + 1) / 2)).slice(-3); // 3桁のページ番号に変換
        var saveFile = new File("~/Desktop/page_" + pageNum + ".jpg");
        var jpgSaveOptions = new JPEGSaveOptions();
        jpgSaveOptions.quality = 12;
        doc.saveAs(saveFile, jpgSaveOptions, true, Extension.LOWERCASE);
        log("Saved page: " + pageNum);

        // 次のページのために新しいドキュメントを作成
        if (i + 2 < csvData.length) {
            doc.close(SaveOptions.DONOTSAVECHANGES); // 前のドキュメントを閉じる
            log("Document closed before creating new one");
            doc = app.documents.add(a4Width, a4Height, 300);
            log("New document created for next page");
        }
    } catch (e) {
        log("Error processing row " + (i + 1) + ": " + e.message);
        alert("Error processing row " + (i + 1) + ": " + e.message);
    }
}

// ドキュメントを閉じる
doc.close(SaveOptions.DONOTSAVECHANGES);
log("Document closed");

logFile.close();
alert("Process complete");

// テキストを配置する関数
function placeText(content, x, y, fontSize, width, bold) {
    try {
        if (!content || (typeof content === "string" && content.replace(/^\s+|\s+$/g, '') === "")) return 0; // コンテンツが空の場合はスキップ
        var textLayer = doc.artLayers.add();
        textLayer.kind = LayerKind.TEXT;
        textLayer.textItem.kind = TextType.PARAGRAPHTEXT; // 段落テキストを設定
        textLayer.textItem.position = [UnitValue(x, "mm"), UnitValue(y, "mm")];
        textLayer.textItem.size = fontSize;
        textLayer.textItem.width = UnitValue(width, "mm");
        textLayer.textItem.height = UnitValue(50, "mm"); // 高さを適切に設定
        textLayer.textItem.contents = content;
        textLayer.textItem.fauxBold = bold;
        textLayer.textItem.justification = Justification.LEFT;

        // フォントをApple Color Emojiに設定
        textLayer.textItem.font = "AppleColorEmoji";

        // レイヤーをアクティブにしてバウンディングボックスを取得
        app.activeDocument.activeLayer = textLayer;
        var textBounds = textLayer.bounds;
        var textHeight = textBounds[3].as("mm") - textBounds[1].as("mm");
        return textHeight;
    } catch (e) {
        log("Error placing text: " + e.message);
        alert("Error placing text: " + e.message);
        throw e;  // エラーを再スローして上位でキャッチする
    }
}

// 画像を配置する関数
function placeImage(imagePath, x, y, maxSize) {
    try {
        log("Placing image: " + imagePath);
        var imageFile = new File(imagePath);
        if (!imageFile.exists) {
            log("Image file not found: " + imagePath);
            alert("Image file not found: " + imagePath);
            return;
        }

        var imageDoc = app.open(imageFile);

        // ドキュメントの解像度を取得
        var docResolution = imageDoc.resolution;

        // 画像のサイズを取得
        var width = imageDoc.width.as("px");
        var height = imageDoc.height.as("px");

        // 最大サイズをピクセルに変換
        var maxSizePixels = maxSize * (docResolution / 25.4);

        // 画像の長辺をリサイズ
        if (width > height) {
            imageDoc.resizeImage(UnitValue(maxSizePixels, "px"), null, null, ResampleMethod.BICUBIC);
        } else {
            imageDoc.resizeImage(null, UnitValue(maxSizePixels, "px"), null, ResampleMethod.BICUBIC);
        }

        imageDoc.selection.selectAll();
        imageDoc.selection.copy();
        imageDoc.close(SaveOptions.DONOTSAVECHANGES);

        doc.paste();
        var pastedLayer = doc.activeLayer;

        // 画像の位置を調整
        var newWidth = pastedLayer.bounds[2].as("mm") - pastedLayer.bounds[0].as("mm");
        var newHeight = pastedLayer.bounds[3].as("mm") - pastedLayer.bounds[1].as("mm");

        pastedLayer.translate(UnitValue(x, "mm") - pastedLayer.bounds[0].as("mm"), UnitValue(y, "mm") - pastedLayer.bounds[1].as("mm"));
        log("Image placed successfully: " + imagePath);
    } catch (e) {
        log("Error placing image: " + e.message);
        alert("Error placing image: " + e.message);
        throw e;  // エラーを再スローして上位でキャッチする
    }
}
