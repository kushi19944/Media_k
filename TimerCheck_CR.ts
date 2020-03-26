import RPA from 'ts-rpa';
// 指定された時間と一致した行を 作業対象に変更するプログラム

// ＊＊＊＊＊＊＊流用時の変更ポイント＊＊＊＊＊＊＊
// スプレッドシートID
const SSID = process.env.Status_CP_SheetID;
// スプレッドシート名
const SSName1 = 'CR_時間指定';
// ＊＊＊＊＊＊＊流用時の変更ポイント＊＊＊＊＊＊＊

async function Start() {
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `CR_時間指定!J8:K500`
  });
  for (let i in FirstData) {
    FirstData[i].push(`${Number(i) + 8}`);
  }

  await RPA.Logger.info(FirstData);
  var dt = new Date();
  var y = dt.getFullYear();
  var m = ('00' + (dt.getMonth() + 1)).slice(-2);
  var d = ('00' + dt.getDate()).slice(-2);
  var NowHours = dt.getHours();
  var result = y + '-' + m + '-' + d;
  var result2 = m + d;
  await RPA.Logger.info('現在:' + result2 + '日 ' + NowHours + '時');
  const a = FirstData[0][0];
  const DataList = [];
  for (let i in FirstData) {
    await Replace(DataList, FirstData[i]);
    if (DataList[0] == result2) {
      if (DataList[1] == NowHours) {
        await RPA.Logger.info(`日時一致 ${DataList[2]} 行目`);
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `CR_時間指定!A${DataList[2]}:A${DataList[2]}`,
          values: [[`作業対象`]]
        });
        await RPA.sleep(500);
      }
    }
  }
}

Start();

// 時間の文字を整えて、List に日付データを格納する
async function Replace(DataList, SheetData) {
  if (SheetData[0].includes('/') == true) {
    if (SheetData[0].indexOf('0') == 0) {
      const ReplaceText = SheetData[0].replace('/', '');
      DataList[0] = ReplaceText;
    }
    // 文字の先頭に 0 が付いていなければ追加してリストに格納する
    if (SheetData[0].indexOf('0') != 0) {
      const ReplaceText = SheetData[0].replace('/', '');
      DataList[0] = '0' + ReplaceText;
    }
  }
  if (SheetData[0].includes('-') == true) {
    if (SheetData[0].indexOf('0') == 0) {
      const ReplaceText = SheetData[0].replace('-', '');
      DataList[0] = ReplaceText;
    }
    // 文字の先頭に 0 が付いていなければ追加してリストに格納する
    if (SheetData[0].indexOf('0') != 0) {
      const ReplaceText = SheetData[0].replace('-', '');
      DataList[0] = '0' + ReplaceText;
    }
  }
  DataList[1] = SheetData[1];
  DataList[2] = SheetData[2];
}
