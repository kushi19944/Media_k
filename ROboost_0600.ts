import RPA from 'ts-rpa';

// ＊＊＊＊＊＊＊＊＊＊ 流用時の変更ポイント ＊＊＊＊＊＊＊＊＊＊
// スプレッドシートID を記載
const SSID = process.env.ROboost_0000_SheetID;
// スプレッドシート名 を記載
const SSName1 = process.env.ROboost_0600_SheetName;
// Slack へ通知する際の文言 (◯時 や 日中トリガーなど)
const SlackText = '6時';
// Slack の通知オンオフ設定 trueなら通知され, falseなら通知オフ
var SlackFlag = true;
// ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

const AJA_ID = process.env.AJA_CAAD_RPA_ID;
const AJA_PW = process.env.AJA_CAAD_RPA_PW;

// サイバーSlack Bot 通知トークン・チャンネル
const BotToken = process.env.CyberBotToken;
const BotChannel = process.env.CyberBotChannel;
// スプレッドシートから読み込む行数を記載する
const StartRow = 8;
const LastRow = 1000;

async function Start() {
  // 実行前にダウンロードフォルダを全て削除する
  await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
  // デバッグログを最小限にする
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  const firstSheetData = [];
  const SheetWorkingRow = [];
  const LoopFlag = ['true'];
  // Slackの通知用フラッグ goodなら完了報告。errorならエラー報告を行う
  const WorkStatus = ['good'];
  // リスト100件表示させるフラグ
  const List100Flag = [false];
  try {
    await SlackPost(`インフィード入札額 ${SlackText}調整 開始します`);
    await AJALogin();
    while (0 == 0) {
      await RPA.Logger.info('＊＊＊開始＊＊＊');
      LoopFlag[0] = 'false';
      // 作業する行のデータを取得
      await ReadSheet(firstSheetData, LoopFlag, SheetWorkingRow);
      // RPAフラグ・管理ツールURL・ADGID・調整入札値
      const SheetData = firstSheetData[0];
      if (LoopFlag[0] == 'true') {
        await TabCreate();
        const PageStatus = ['good'];
        await PageMoving(SheetData, SheetWorkingRow, PageStatus, List100Flag);
        if (PageStatus[0] == 'good') {
          await TargetInputSelect(SheetData, SheetWorkingRow);
        }
        await TabClose();
      }
      if (LoopFlag[0] == 'false') {
        await RPA.Logger.info('全ての行の処理完了しました');
        await RPA.Logger.info('ループ処理ブレイク');
        break;
      }
    }
  } catch {
    await RPA.Logger.info('エラー発生 Slackにてエラー通知します');
    WorkStatus[0] = 'error';
    await RPA.WebBrowser.takeScreenshot();
    await RPA.Logger.info('スクリーンショット撮ってブラウザ終了します');
  } finally {
    await RPA.WebBrowser.quit();
    if (WorkStatus[0] == 'good') {
      // 問題なければ完了報告を行う
      await SlackPost(
        `インフィード入札額 ${SlackText}調整 問題なく完了しました`
      );
    }
    if (WorkStatus[0] == 'error') {
      // 問題があればエラー報告を行う
      await SlackPost(
        `インフィード入札額 ${SlackText}調整 エラーが発生しました\n@kushi_makoto 確認してください`
      );
    }
  }
}

Start();

async function SlackPost(Text) {
  // SlackFlagが trueなら Slackにメッセージ投稿
  if (SlackFlag == true) {
    await RPA.Slack.chat.postMessage({
      channel: BotChannel,
      token: BotToken,
      text: `${Text}`,
      icon_emoji: ':snowman:',
      username: 'p1',
    });
  }
}

// Tabを作成し,2番目に切り替える
async function TabCreate() {
  await RPA.WebBrowser.driver.executeScript(`window.open('')`);
  await RPA.sleep(200);
  const tab = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(tab[1]);
}

// Tabを閉じて1番目に切り替える
async function TabClose() {
  const tab = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(tab[0]);
  await RPA.WebBrowser.closeWindow(tab[1]);
}

async function ReadSheet(SheetData, LoopFlag, SheetWorkingRow) {
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A${String(StartRow)}:D${String(LastRow)}`,
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
        SheetData[0] = FirstData[i];
        LoopFlag[0] = 'true';
        const Row = Number(i) + Number(StartRow);
        SheetWorkingRow[0] = Row;
        await SetValues_Function(SheetWorkingRow, `作業中`);
        break;
      }
    }
  }
}

async function AJALogin() {
  await RPA.WebBrowser.get(process.env.ROboost_URL);
  const IDInput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'mailAddress' }),
    5000
  );
  const PWInput = await RPA.WebBrowser.findElementById('password');
  const LoginButton = await RPA.WebBrowser.findElementById('submit');
  await RPA.WebBrowser.sendKeys(IDInput, [AJA_ID]);
  await RPA.WebBrowser.sendKeys(PWInput, [AJA_PW]);
  await RPA.WebBrowser.mouseClick(LoginButton);
  for (let i = 0; i < 100; i++) {
    try {
      await RPA.sleep(500);
      const UserAria = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({ className: 'user-area' }),
        5000
      );
      const UserAriaText = await UserAria.getText();
      if (UserAriaText.indexOf(AJA_ID) >= 0) {
        await RPA.Logger.info('ログインできました');
        break;
      }
    } catch {}
    if (i == 90) {
      await SlackPost(
        `インフィード入札額 ${SlackText}調整 ログインエラー. IDパスワードの確認をお願いします`
      );
      await RPA.sleep(1000);
      process.exit(0);
    }
  }
}

async function PageMoving(SheetData, SheetWorkingRow, PageStatus, List100Flag) {
  // ログイン時に1回だけ100件 をクリックする
  if (List100Flag[0] == false) {
    List100Flag[0] = true;
    await RPA.WebBrowser.get(SheetData[1]);
    await RPA.sleep(1000);
    const CampaignID = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        css: '#listTableCampaign > tbody > tr:nth-child(1) > td:nth-child(2)',
      }),
      5000
    );
    const PullSelect = await RPA.WebBrowser.findElementByCSSSelector(
      `#main > article > div.contents.ng-scope > section > div:nth-child(7) > div > div:nth-child(2) > div > dl > dd > select > option:nth-child(4)`
    );
    await PullSelect.click();
    await RPA.sleep(4000);
    await RPA.Logger.info('表示数を100に変更します');
  }
  const NewURL = SheetData[1].replace('campaign?', 'campaign/adgroup?');
  await RPA.WebBrowser.get(NewURL);
}

async function TargetInputSelect(SheetData, SheetWorkingRow) {
  for (let v = 0; v < 20; v++) {
    const Allbrake = ['false'];
    for (let i = 0; i < 5; i++) {
      try {
        // IDが出現するまで待機 (4回リトライ)
        const AdgID = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            css:
              '#listTableAdGroup > tbody > tr:nth-child(1) > td:nth-child(3)',
          }),
          5000
        );
      } catch {
        // 4回リトライしてもID出ない場合はスキップ
        if (i == 4) {
          await SetValues_Function(SheetWorkingRow, `ID表示されませんでした`);
          // 親ループもブレイクさせる
          Allbrake[0] = 'true';
          break;
        }
      }
    }
    for (let NewNumber = 1; NewNumber < 101; NewNumber++) {
      try {
        var ID = await RPA.WebBrowser.findElementByCSSSelector(
          `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`
        );
      } catch {
        await SetValues_Function(SheetWorkingRow, `ID不一致`);
        // 親ループもブレイクさせる
        Allbrake[0] = 'true';
        break;
      }
      const IDText = await ID.getText();
      if (IDText == SheetData[2]) {
        // 親ループもブレイクさせる
        Allbrake[0] = 'true';
        await RPA.Logger.info('【入札調整】ID一致 ' + IDText);
        await RPA.WebBrowser.scrollTo({
          selector: `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`,
        });
        //await RPA.Logger.info('【入札調整】入力したと想定');
        await RPA.sleep(300);
        const YenClick = await RPA.WebBrowser.findElementByCSSSelector(
          `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td.numeric.auto-bid-status > div > editable-box > form > div > div.numeric.ng-scope > a`
        );
        await RPA.WebBrowser.mouseClick(YenClick);
        await RPA.sleep(700);
        const YenInput = await RPA.WebBrowser.findElementByCSSSelector(
          `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td.numeric.auto-bid-status > div > editable-box > form > div > input`
        );
        await RPA.Logger.info('【入札調整】値入力');
        await YenInput.clear();
        await RPA.sleep(300);
        await RPA.WebBrowser.sendKeys(YenInput, [SheetData[3]]);
        await RPA.WebBrowser.sendKeys(YenInput, [RPA.WebBrowser.Key.ENTER]);
        await RPA.sleep(800);
        await SetValues_Function(SheetWorkingRow, `完了`);
        break;
      }
      // 100件毎に検索してIDが一致しなければ次のページへいく
      if (NewNumber == 100) {
        await RPA.WebBrowser.driver.executeScript(
          `document.getElementsByClassName('pagination-next ng-scope')[0].children[0].click()`
        );
        await RPA.Logger.info('【入札調整】次のページでID検索');
        await RPA.sleep(700);
        break;
      }
    }
    if (Allbrake[0] == 'true') {
      // 親ループ処理 ブレイク
      break;
    }
    if (v == 19) {
      await SetValues_Function(SheetWorkingRow, `ID不一致`);
    }
  }
}

// スプレッドシートに貼り付ける関数
async function SetValues_Function(SheetWorkingRow, Text) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A${SheetWorkingRow[0]}:A${SheetWorkingRow[0]}`,
    values: [[`${Text}`]],
  });
}
