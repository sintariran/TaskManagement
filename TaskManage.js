const MAX_TOKEN_NUM = 2048;

function handleUserPrompt(prompt) {
    // 1. 現在のシートフォーマットとデータを取得
    let sheetFormat = getSheetFormat();
    let sheetData = getSheetData();
    
    // 2. APIに送信するプロンプトを構築
    let apiPrompt = buildApiPrompt(sheetFormat, sheetData, prompt);
    
    // 3. APIを呼び出し、応答を取得
    let apiResponse = callOpenAiApi(apiPrompt, 'gpt-4o');
    
    // 4. 応答に基づいてシートを更新
    if (apiResponse) {
        updateSheetBasedOnApiResponse(apiResponse);
    } else {
        Logger.log('API応答がありません。');
        return null;
    }

    // APIの応答をパースして返す
    try {
        Logger.log('API応答のパースを開始します。');
        let jsonResponse = apiResponse.match(/```json\s*([\s\S]*?)\s*```/);
        Logger.log('正規表現によるマッチング完了: ' + JSON.stringify(jsonResponse));
        
        let jsonString;
        if (jsonResponse && jsonResponse[1]) {
            Logger.log('JSON部分の抽出に成功しました。');
            // エスケープされた文字列を正しいJSON形式に変換
            jsonString = jsonResponse[1].replace(/\\n/g, '\n').replace(/\\"/g, '"').replace(/\\\\/g, '\\');
        } else {
            Logger.log('API応答のJSON部分が見つかりません。直接パースを試みます。');
            let response = JSON.parse(apiResponse);
            jsonString = response.choices[0].message.content;
        }

        Logger.log('変換後のJSON文字列: ' + jsonString);
        let parsedResponse = JSON.parse(jsonString);
        Logger.log('JSONパースに成功しました。');
        return parsedResponse;
    } catch (e) {
        Logger.log('API応答のパースエラー: ' + e.message);
        return null;
    }
}
function updateSheetBasedOnApiResponse(responseText) {
    try {
        Logger.log('API応答のテキスト: ' + responseText);
        let response = JSON.parse(responseText);
        Logger.log('パースされたAPI応答: ' + JSON.stringify(response));
        
        if (!response.choices || !response.choices[0].message.content) {
            Logger.log('API応答の形式が不正です。');
            return;
        }
        
        let actionsContent = response.choices[0].message.content;
        Logger.log('API応答のcontent部分: ' + actionsContent);

        // ```json と ``` を削除して有効なJSONに変換
        actionsContent = actionsContent.replace(/```json/g, '').replace(/```/g, '').trim();
        Logger.log('整形後のJSON文字列: ' + actionsContent);
        
        // JSONとしてパース
        let actionsObject;
        try {
            actionsObject = JSON.parse(actionsContent);
            Logger.log('パースされたアクションオブジェクト: ' + JSON.stringify(actionsObject));
        } catch (e) {
            Logger.log('部分的なJSONパースエラー: ' + e.message);
            actionsContent += ']}'; // 応答が途中で切れている場合の対処
            actionsObject = JSON.parse(actionsContent);
            Logger.log('修正後のパースされたアクションオブジェクト: ' + JSON.stringify(actionsObject));
        }

        if (!Array.isArray(actionsObject.actions)) {
            Logger.log('API応答のactionsが配列ではありません。');
            return;
        }
        
        let actions = actionsObject.actions;  // アクション配列を取得
        Logger.log('取得したアクション配列: ' + JSON.stringify(actions));
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

        actions.forEach(action => {
            Logger.log('処理中のアクション: ' + JSON.stringify(action));
            switch (action.action) {
                case 'update':
                    updateRowByTaskId(sheet, action.row, action.column, action.value);
                    Logger.log(`更新アクション: 行 ${action.row}, 列 ${action.column}, 新しい値: ${action.value}`);
                    break;
                case 'delete':
                    deleteRow(sheet, action.row);
                    Logger.log(`削除アクション: 行 ${action.row}`);
                    break;
                case 'add':
                    addRow(sheet, action.values);
                    Logger.log(`追加アクション: ${JSON.stringify(action.values)}`);
                    break;
                default:
                    Logger.log(`不明なアクション: ${JSON.stringify(action)}`);
            }
        });
    } catch (e) {
        Logger.log(`JSONパースエラー: ${e.message}`);
        Logger.log(`API応答内容: ${responseText}`);
    }
}

function getSheetFormat() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers;
}

function getSheetData() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    return sheet.getDataRange().getValues();
}

function buildApiPrompt(sheetFormat, sheetData, userPrompt) {
    let prompt = `
        現在のシートフォーマットは以下の通りです：${JSON.stringify(sheetFormat)}。
        現在のシートデータは以下の通りです：${JSON.stringify(sheetData)}。
        ユーザープロンプト：${userPrompt}
        変更が必要な部分のみを特定し、その変更内容だけを以下の形式で返してください。
        応答フォーマットは以下の通りです：
        {
            "actions": [
                { "action": "update", "row": 2, "column": "締め切り", "value": "2024年6月1日" },
                { "action": "delete", "row": 4 },
                { "action": "add", "values": ["新しいタスク", "詳細", "優先度", "2024年6月10日"] }
            ]
        }
    `;
    Logger.log(`API送信内容: ${prompt}`);
    return prompt;
}

function callOpenAiApi(prompt, modelName) {
    let apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    try {
        let response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
            method: "post",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
            },
            payload: JSON.stringify({
                model: modelName,
                max_tokens: MAX_TOKEN_NUM,
                temperature: 0,
                messages: [{ role: "user", content: prompt }]
            })
        });
        let responseText = response.getContentText();
        Logger.log(`API応答内容: ${responseText}`);
        return responseText;
    } catch (e) {
        Logger.log(`API呼び出しエラー: ${e.message}`);
        return null;
    }
}

function updateRow(sheet, rowIndex, columnName, newValue) {
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let columnIndex = headers.indexOf(columnName) + 1;
    if (columnIndex > 0) {
        sheet.getRange(rowIndex, columnIndex).setValue(newValue);
    } else {
        Logger.log(`Column ${columnName} not found. Available columns: ${headers.join(', ')}`);
    }
}

function deleteRow(sheet, rowIndex) {
    sheet.deleteRow(rowIndex);
}

function addRow(sheet, values) {
    sheet.appendRow(values);
}

// カスタムメニューを作成する関数
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('タスク管理')
        .addItem('タスク管理ツールを開く', 'showSidebar')
        .addToUi();
}

// サイドバーを表示する関数
function showSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('index')
        .setTitle('タスク管理ツール');
    SpreadsheetApp.getUi().showSidebar(html);
}

function getTasksHtml() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let data = sheet.getDataRange().getValues();
    let tasks = data.slice(1); // ヘッダー行を除く

    let taskListHtml = tasks.map(task => {
        return `<li class="task-item">${task.join(' - ')}</li>`;
    }).join('');

    return taskListHtml;
}

function updateRowByTaskId(sheet, rowIndex, columnName, newValue) {
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let columnIndex = headers.indexOf(columnName) + 1;

    if (columnIndex <= 0) {
        Logger.log(`Column ${columnName} not found. Available columns: ${headers.join(', ')}`);
        return;
    }

    // 行数を直接使用して更新
    sheet.getRange(rowIndex, columnIndex).setValue(newValue);
    Logger.log(`更新アクション: 行 ${rowIndex}, 列 ${columnName}, 新しい値: ${newValue}`);
}

function doGet() {
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('タスク管理ツール')
        .setWidth(400);
}