const MAX_TOKEN_NUM = 150;

function handleUserPrompt(prompt) {
    // 1. 現在のシートフォーマットとデータを取得
    let sheetFormat = getSheetFormat();
    let sheetData = getSheetData();
    
    // 2. APIに送信するプロンプトを構築
    let apiPrompt = buildApiPrompt(sheetFormat, sheetData, prompt);
    
    // 3. APIを呼び出し、応答を取得
    let apiResponse = callOpenAiApi(apiPrompt, 'gpt-4o');
    
    // 4. 応答に基づいてシートを更新
    updateSheetBasedOnApiResponse(apiResponse);
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

function updateSheetBasedOnApiResponse(responseText) {
    if (!responseText) {
        Logger.log('API応答がnullです。');
        return;
    }

    try {
        let response = JSON.parse(responseText);
        if (!response.choices || !response.choices[0].message.content) {
            Logger.log('API応答の形式が不正です。');
            return;
        }
        
        let actionsContent = response.choices[0].message.content;

        // ```json と ``` を削除して有効なJSONに変換
        actionsContent = actionsContent.replace(/```json/g, '').replace(/```/g, '').trim();
        
        let actionsObject = JSON.parse(actionsContent);
        let actions = actionsObject.actions;  // アクション配列を取得
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

        actions.forEach(action => {
            switch (action.action) {
                case 'update':
                    updateRow(sheet, action.row, action.column, action.value);
                    break;
                case 'delete':
                    deleteRow(sheet, action.row);
                    break;
                case 'add':
                    addRow(sheet, action.values);
                    break;
            }
        });
    } catch (e) {
        Logger.log(`JSONパースエラー: ${e.message}`);
        Logger.log(`API応答内容: ${responseText}`);
    }
}

function updateRow(sheet, rowIndex, columnName, newValue) {
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let columnIndex = headers.indexOf(columnName) + 1;
    if (columnIndex > 0) {
        sheet.getRange(rowIndex, columnIndex).setValue(newValue);
    } else {
        Logger.log(`Column ${columnName} not found`);
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
