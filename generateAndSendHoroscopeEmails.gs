function generateAndSendHoroscopeEmails() {
  // スプレッドシートとシートの取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheetByName("Main");
  const templateSheet = spreadsheet.getSheetByName("Templates");
  
  // テンプレートの取得
  const horoscopeTemplates = templateSheet.getRange("A1:A").getValues().flat().filter(Boolean);
  const actionTemplates = templateSheet.getRange("B1:B").getValues().flat().filter(Boolean);

  // メインシートのデータ取得
  const data = mainSheet.getDataRange().getValues();

  // 今日の運勢とおすすめ行動をランダムに割り当て、メール送信
  for (let i = 1; i < data.length; i++) { // 1行目はヘッダーとしてスキップ
    const [name, birthday, hobby, , , email] = data[i];
    if (!name || !email) continue; // 名前またはメールアドレスが空の場合スキップ

    // ランダムに選択
    const horoscope = horoscopeTemplates[Math.floor(Math.random() * horoscopeTemplates.length)];
    const action = actionTemplates[Math.floor(Math.random() * actionTemplates.length)];

    // メインシートを更新
    mainSheet.getRange(i + 1, 4).setValue(horoscope); // 運勢をD列に記入
    mainSheet.getRange(i + 1, 5).setValue(`${action} (${hobby ? hobby + "に関連するとさらに良い！" : ""})`); // 行動をE列に記入

    // メール送信
    const subject = `今日の運勢とおすすめ行動 - ${name}さんへ`;
    const message = `${name}さん、\n\n今日の運勢: ${horoscope}\nおすすめの行動: ${action}\n\n素敵な一日を！`;

    GmailApp.sendEmail(email, subject, message);
  }
}
