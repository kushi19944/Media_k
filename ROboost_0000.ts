import RPA from 'ts-rpa';

// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.ROboost_0000_SheetID;
const SSName1 = process.env.ROboost_0000_SheetName;
// サイバーSlack Bot 通知トークン・チャンネル
const BotToken = process.env.CyberBotToken;
const BotChannel = process.env.CyberBotChannel;
// Slack へ通知する際の文言
const SlackText = '0時';
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
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  const firstSheetData = [];
  const SheetWorkingRow = [];
  const LoopFlag = ['true'];
  // Slackの通知用フラッグ goodなら完了報告。errorならエラー報告を行う
  const WorkStatus = ['good'];
  try {
    await SlackStart();
    await AJALogin();
    while (0 == 0) {
      LoopFlag[0] = 'false';
      // 作業する行のデータを取得
      await ReadSheet(firstSheetData, LoopFlag, SheetWorkingRow);
      // RPAフラグ・管理ツールURL・ADGID・調整入札値
      const SheetData = firstSheetData[0];
      await RPA.Logger.info(SheetData);
      if (LoopFlag[0] == 'true') {
        await TabCreate();
        const PageStatus = ['good'];
        await PageMoveing(SheetData, SheetWorkingRow, PageStatus);
        if (PageStatus[0] == 'good') {
          await TargetInputSelect(SheetData, SheetWorkingRow);
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
  } catch {
    await RPA.Logger.info('エラー発生 Slackにてエラー通知します');
    WorkStatus[0] = 'error';
    await RPA.WebBrowser.takeScreenshot();
    await RPA.Logger.info('スクリーンショット撮ってブラウザ終了します');
  } finally {
    await RPA.WebBrowser.quit();
    if (WorkStatus[0] == 'good') {
      // 問題なければ完了報告を行う
      await RPA.Slack.chat.postMessage({
        channel: BotChannel,
        token: BotToken,
        text: `インフィード入札額 ${SlackText}調整 問題なく完了しました`,
        icon_emoji: ':snowman:',
        username: 'p1'
      });
    }
    if (WorkStatus[0] == 'error') {
      // 問題があればエラー報告を行う
      await RPA.Slack.chat.postMessage({
        channel: BotChannel,
        token: BotToken,
        text: `インフィード入札額 ${SlackText}調整 エラーが発生しました\n@kushi_makoto 確認してください`,
        icon_emoji: ':snowman:',
        username: 'p1'
      });
    }
  }
}

Start();

async function SlackStart() {
  // 作業開始時にSlackへ通知する
  await RPA.Slack.chat.postMessage({
    channel: BotChannel,
    token: BotToken,
    text: `インフィード入札額 ${SlackText}調整 開始します`,
    icon_emoji: ':snowman:',
    username: 'p1'
  });
}

// Tabを作成し,2番目に切り替える
async function TabCreate() {
  await RPA.WebBrowser.driver.executeScript(`window.open('')`);
  await RPA.sleep(200);
  const tab = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.Logger.info(tab);
  await RPA.WebBrowser.switchToWindow(tab[1]);
  await RPA.Logger.info('新規タブに切り替えます');
  await RPA.sleep(500);
}

// Tabを閉じて1番目に切り替える
async function TabClose() {
  const tab = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.Logger.info(tab);
  await RPA.WebBrowser.switchToWindow(tab[0]);
  await RPA.WebBrowser.closeWindow(tab[1]);
  await RPA.sleep(500);
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
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName1}!A${SheetWorkingRow[0]}:A${SheetWorkingRow[0]}`,
          values: [['作業中']]
        });
        break;
      }
    }
  }
}

async function AJALogin() {
  await RPA.WebBrowser.get('https://agency.aja.fm/#/account');
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

async function PageMoveing(SheetData, SheetWorkingRow, PageStatus) {
  await RPA.Logger.info(`このURLに飛びます ${SheetData[1]}`);
  await RPA.WebBrowser.get(SheetData[1]);
  while (0 == 0) {
    try {
      await RPA.sleep(500);
      const UserAria = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({ className: 'user-area' }),
        8000
      );
      const UserAriaText = await UserAria.getText();
      await RPA.Logger.info(UserAriaText);
      if (UserAriaText.indexOf(process.env.AJA_ROboost_ID) >= 0) {
        await RPA.Logger.info('ユーザーエリア出現しました。次の処理に進みます');
        break;
      }
    } catch {}
  }
  await RPA.sleep(300);
  await RPA.WebBrowser.scrollTo({
    selector:
      '#main > article > div.contents.ng-scope > section > div.list-ui-group.clear > ul.tab'
  });
  await RPA.sleep(200);
  const KoukokuGruop = await RPA.WebBrowser.findElementByCSSSelector(
    '#main > article > div.contents.ng-scope > section > div.list-ui-group.clear > ul.tab > li:nth-child(2)'
  );
  await RPA.WebBrowser.mouseClick(KoukokuGruop);
  await RPA.sleep(500);
  // たまにページが表示されないことがあるため、15秒待って出ない時はスキップする
  try {
    const GenzaiNyuusatsu = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        css: '#listTableAdGroup > thead > tr > th.current_daily_budget'
      }),
      15000
    );
  } catch {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName1}!A${SheetWorkingRow[0]}:A${SheetWorkingRow[0]}`,
      values: [['ページが開けません']]
    });
    PageStatus[0] = 'bad';
    return;
  }
  // 100件表示させる
  const PullSelect = await RPA.WebBrowser.findElementByCSSSelector(
    `#main > article > div.contents.ng-scope > section > div:nth-child(7) > div > div:nth-child(2) > div > dl > dd > select > option:nth-child(4)`
  );
  await PullSelect.click();
  await RPA.sleep(4000);
}

async function TargetInputSelect(SheetData, SheetWorkingRow) {
  const ThisPageURL = await RPA.WebBrowser.getCurrentUrl();
  await RPA.Logger.info(ThisPageURL);
  for (let v = 2; v < 10; v++) {
    const Allbrake = ['false'];
    try {
      for (let NewNumber = 1; NewNumber < 101; NewNumber++) {
        var ID = await RPA.WebBrowser.findElementByCSSSelector(
          `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`
        );
        const IDText = await ID.getText();
        if (IDText == SheetData[2]) {
          // 親ループもブレイクさせる
          Allbrake[0] = 'true';
          await RPA.Logger.info('ID一致しました');
          await RPA.WebBrowser.scrollTo({
            selector: `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`
          });
          await RPA.sleep(300);
          const YenClick = await RPA.WebBrowser.findElementByCSSSelector(
            `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td.numeric.auto-bid-status > div > editable-box > form > div > div.numeric.ng-scope > a`
          );
          await RPA.WebBrowser.mouseClick(YenClick);
          await RPA.sleep(700);
          const YenInput = await RPA.WebBrowser.findElementByCSSSelector(
            `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td.numeric.auto-bid-status > div > editable-box > form > div > input`
          );
          await RPA.Logger.info('調整入札値 入力します');
          await YenInput.clear();
          await RPA.sleep(300);
          await RPA.WebBrowser.sendKeys(YenInput, [SheetData[3]]);
          await RPA.WebBrowser.sendKeys(YenInput, [RPA.WebBrowser.Key.ENTER]);
          await RPA.sleep(800);
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName1}!A${SheetWorkingRow[0]}:A${SheetWorkingRow[0]}`,
            values: [['完了']]
          });
          break;
        }
        // 100件毎に検索してIDが一致しなければ次のページへいく
        if (NewNumber == 100) {
          const NextPageURL = ThisPageURL + `&page=${v}`;
          await RPA.Logger.info('次のページへ移動してID検索します');
          await RPA.WebBrowser.get(NextPageURL);
          await RPA.sleep(7000);
          break;
        }
      }
    } catch {
      // IDが見つからない時は A列をエラー表示に変更
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName1}!A${SheetWorkingRow[0]}:A${SheetWorkingRow[0]}`,
        values: [['ID不一致']]
      });
      break;
    }
    if (Allbrake[0] == 'true') {
      await RPA.Logger.info('ID一致したので全てブレイクします');
      break;
    }
  }
}
