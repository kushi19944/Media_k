import RPA from 'ts-rpa';

// ＊＊＊＊＊＊＊流用時の変更ポイント＊＊＊＊＊＊＊
// スプレッドシートID
const SSID = process.env.Status_CP_SheetID;
// スプレッドシート名
const SSName1 = process.env.Status_CR_SheetName;
// ＊＊＊＊＊＊＊流用時の変更ポイント＊＊＊＊＊＊＊

// スプレッドシートから読み込む行数を記載する
const StartRow = 8;
const LastRow = 1000;
// AJA管理ツールの 目的のタブ をクリックする
async function TargetTabClick() {
  // クリエイティブ
  const TargetTab = await RPA.WebBrowser.findElementByCSSSelector(
    '#main > article > div.contents.ng-scope > section > div.list-ui-group.clear > ul.tab > li:nth-child(4)'
  );
  await RPA.WebBrowser.mouseClick(TargetTab);
}

async function Start() {
  // 実行前にダウンロードフォルダを全て削除する
  await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
  // デバッグログを最小限にする
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  try {
    await Work();
  } catch {
    // エラー発生時の処理
    //const DOM = await RPA.WebBrowser.driver.getPageSource();
    //await RPA.Logger.info(DOM);
    await RPA.WebBrowser.takeScreenshot();
    await RPA.Logger.info(
      'エラー出現.スクリーンショット撮ってブラウザ終了します'
    );
  } finally {
    await RPA.WebBrowser.quit();
  }
}

Start();

async function Work() {
  const firstSheetData = [];
  const SheetWorkingRow = [];
  const LoopFlag = ['true'];
  await AJALogin();
  while (0 == 0) {
    LoopFlag[0] = 'false';
    // 作業する行のデータを取得
    await ReadSheet(firstSheetData, LoopFlag, SheetWorkingRow);
    // RPAフラグ・管理ツールURL・ID・有効or無効
    const SheetData = firstSheetData[0];
    await RPA.Logger.info(SheetData);
    if (LoopFlag[0] == 'true') {
      await TabCreate();
      const PageStatus = ['good'];
      await PageMoveing(SheetData, SheetWorkingRow, PageStatus);
      if (PageStatus[0] == 'good') {
        await StatusChange(SheetData, SheetWorkingRow);
      }
      await RPA.sleep(300);
      TabClose();
    }
    if (LoopFlag[0] == 'false') {
      await RPA.Logger.info('全ての行の処理完了しました');
      await RPA.Logger.info('ループ処理ブレイク');
      break;
    }
  }
}

async function AJALogin() {
  await RPA.WebBrowser.get(process.env.Status_URL);
  const IDInput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'mailAddress' }),
    5000
  );
  const PWInput = await RPA.WebBrowser.findElementById('password');
  const LoginButton = await RPA.WebBrowser.findElementById('submit');
  await RPA.WebBrowser.sendKeys(IDInput, [process.env.AJA_ROboost_ID]);
  await RPA.WebBrowser.sendKeys(PWInput, [process.env.AJA_ROboost_PW]);
  await RPA.WebBrowser.mouseClick(LoginButton);
  while (0 == 0) {
    try {
      await RPA.sleep(500);
      const UserAria = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({ className: 'user-area' }),
        5000
      );
      const UserAriaText = await UserAria.getText();
      await RPA.Logger.info(UserAriaText);
      if (UserAriaText.indexOf(process.env.AJA_ROboost_ID) >= 0) {
        await RPA.Logger.info('ログインできました');
        break;
      }
    } catch {}
  }
}

async function ReadSheet(SheetData, LoopFlag, SheetWorkingRow) {
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A${String(StartRow)}:D${String(LastRow)}`
  });
  // B列にURL と D列に調整入札値が 入っていてかつ、A列が空白 の行だけ取得する
  for (let i in FirstData) {
    if (FirstData[i][1] != '') {
      if (FirstData[i][2] == '') {
        continue;
      }
      if (FirstData[i][2] == undefined) {
        continue;
      }
      if (FirstData[i][3] == '') {
        continue;
      }
      if (FirstData[i][3] == undefined) {
        continue;
      }
      if (FirstData[i][0] == '') {
        await RPA.Logger.info(FirstData[i]);
        SheetData[0] = FirstData[i];
        LoopFlag[0] = 'true';
        const Row = Number(i) + Number(StartRow);
        SheetWorkingRow[0] = Row;
        await PasteSheet('作業中', SheetWorkingRow);
        break;
      }
    }
  }
}

// スプレッドシートに 現状ステータス を記載する
async function PasteSheet(StatusText, SheetWorkingRow) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A${SheetWorkingRow[0]}:A${SheetWorkingRow[0]}`,
    values: [[StatusText]]
  });
}

// Tabを作成し,2番目に切り替える
async function TabCreate() {
  await RPA.WebBrowser.driver.executeScript(`window.open('')`);
  await RPA.sleep(200);
  const tab = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(tab[1]);
  await RPA.Logger.info('新規タブに切り替えます');
  await RPA.sleep(500);
}

// Tabを閉じて1番目に切り替える
async function TabClose() {
  const tab = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(tab[0]);
  await RPA.WebBrowser.closeWindow(tab[1]);
  await RPA.sleep(500);
}

// 指定のページ(タブ)に移動する
async function PageMoveing(SheetData, SheetWorkingRow, PageStatus) {
  await RPA.WebBrowser.get(SheetData[1]);
  while (0 == 0) {
    try {
      await RPA.sleep(500);
      const UserAria = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({ className: 'user-area' }),
        8000
      );
      const UserAriaText = await UserAria.getText();
      if (UserAriaText.indexOf(process.env.AJA_ROboost_ID) >= 0) {
        await RPA.Logger.info('ユーザーエリア出現しました。次の処理に進みます');
        break;
      }
    } catch {}
  }
  await RPA.sleep(300);
  // タブ の位置まで スクロールする
  await RPA.WebBrowser.scrollTo({
    selector:
      '#main > article > div.contents.ng-scope > section > div.list-ui-group.clear > ul.tab'
  });
  await RPA.sleep(200);
  // 100件表示させる
  const PullSelect = await RPA.WebBrowser.findElementByCSSSelector(
    `#main > article > div.contents.ng-scope > section > div:nth-child(7) > div > div:nth-child(2) > div > dl > dd > select > option:nth-child(4)`
  );
  await PullSelect.click();
  await RPA.sleep(4000);
  // 目的のタブに直接飛ぶ
  const PageURL = await RPA.WebBrowser.getCurrentUrl();
  const TargetURL = PageURL.replace('campaign?', 'campaign/creative?');
  await RPA.WebBrowser.get(TargetURL);
  await RPA.sleep(5000);
  // たまにページが表示されないことがあるため、60秒待って出ない時はスキップする
  try {
    const ID_no1 = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        css: '#listTableCreative > tbody > tr:nth-child(1) > td:nth-child(3)'
      }),
      60000
    );
    await RPA.Logger.info('ID出現しました');
  } catch {
    PageStatus[0] = 'bad';
    await PasteSheet('ページが開けません', SheetWorkingRow);
  }
}

// ステータス変更 メインの処理
async function StatusChange(SheetData, SheetWorkingRow) {
  // 次ページの移動に必要なため CurrentURL を取得しておく
  const ThisPageURL = await RPA.WebBrowser.getCurrentUrl();
  await RPA.Logger.info(ThisPageURL);
  for (let v = 2; v < 11; v++) {
    const Allbrake = ['false'];
    for (let NewNumber = 1; NewNumber < 101; NewNumber++) {
      var ID = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          css: `#listTableCreative > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`
        }),
        60000
      );
      const IDText = await ID.getText();
      if (IDText == SheetData[2]) {
        // 親ループもブレイクさせる
        Allbrake[0] = 'true';
        await RPA.Logger.info('ID一致しました');
        await RPA.WebBrowser.scrollTo({
          selector: `#listTableCreative > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`
        });
        await RPA.sleep(200);
        // 一致したIDの 使用広告 をJavaScriptで直接クリックする
        await RPA.WebBrowser.driver.executeScript(
          `document.querySelector('#listTableCreative > tbody > tr:nth-child(${NewNumber}) > td:nth-child(18) > ul > li:nth-child(2) > button').click()`
        );
        await RPA.sleep(1000);
        const Yuukou = await RPA.WebBrowser.findElementByCSSSelector(
          'body > div.modal.fade.ng-isolate-scope.in > div > div > section > div > ul > li > div > label:nth-child(1)'
        );
        const Mukou = await RPA.WebBrowser.findElementByCSSSelector(
          'body > div.modal.fade.ng-isolate-scope.in > div > div > section > div > ul > li > div > label:nth-child(2)'
        );
        const ApplyButton = await RPA.WebBrowser.findElementByCSSSelector(
          'body > div.modal.fade.ng-isolate-scope.in > div > div > section > div > ul > apply-close-button > li:nth-child(1) > button'
        );
        if (SheetData[3] == '有効') {
          await RPA.WebBrowser.mouseClick(Yuukou);
          await RPA.sleep(300);
          try {
            //await RPA.WebBrowser.mouseClick(ApplyButton);
            await PasteSheet('完了', SheetWorkingRow);
            await RPA.sleep(3000);
            break;
          } catch {
            await RPA.Logger.info('適用ボタン 押せませんでした');
            await PasteSheet('ステータス変更なし', SheetWorkingRow);
            break;
          }
        }
        if (SheetData[3] == '無効') {
          await RPA.WebBrowser.mouseClick(Mukou);
          await RPA.sleep(300);
          try {
            //await RPA.WebBrowser.mouseClick(ApplyButton);
            await RPA.Logger.info('適用ボタン　押したと想定');
            await PasteSheet('完了', SheetWorkingRow);
            await RPA.sleep(3000);
            break;
          } catch {
            await RPA.Logger.info('適用ボタン 押せませんでした');
            await PasteSheet('ステータス変更なし', SheetWorkingRow);
            break;
          }
        }
      }
      // 100件毎に検索してIDが一致しなければ次のページへいく
      if (NewNumber == 100) {
        const NextPageURL = ThisPageURL + `&page=${v}`;
        await RPA.Logger.info('次のページへ移動してID検索します');
        await RPA.WebBrowser.get(NextPageURL);
        await RPA.sleep(5000);
        break;
      }
    }
    if (Allbrake[0] == 'true') {
      await RPA.Logger.info('親ループブレイクします');
      break;
    }
    if (v == 10) {
      // IDが見つからない時は A列をエラー表示に変更
      await PasteSheet('ID不一致', SheetWorkingRow);
      break;
    }
  }
}
