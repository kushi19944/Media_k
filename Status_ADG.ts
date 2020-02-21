import RPA from 'ts-rpa';

// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.Status_CP_SheetID;
const SSName1 = process.env.Status_ADG_SheetName;
// サイバーSlack Bot 通知トークン・チャンネル
const BotToken = process.env.CyberBotToken;
const BotChannel = process.env.CyberBotChannel;
// Slack へ通知する際の文言
const SlackText = '広告グループ';
// スプレッドシートから読み込む行数を記載する
const StartRow = 8;
const LastRow = 1000;
// AJA管理ツールの目的の タブを クリックする
async function TargetTabClick() {
  // 広告グループ
  const TargetTab = await RPA.WebBrowser.findElementByCSSSelector(
    '#main > article > div.contents.ng-scope > section > div.list-ui-group.clear > ul.tab > li:nth-child(2)'
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
  // Slackの通知用フラッグ goodなら完了報告。errorならエラー報告を行う
  const WorkStatus = ['good'];
  try {
    await Work();
  } catch {
    // エラー発生時の処理
    await RPA.Logger.info('エラー発生 Slackにてエラー通知します');
    WorkStatus[0] = 'error';
    const DOM = await RPA.WebBrowser.driver.getPageSource();
    await RPA.Logger.info(DOM);
    await RPA.WebBrowser.takeScreenshot();
    await RPA.Logger.info(
      'エラー出現.スクリーンショット撮ってブラウザ終了します'
    );
  } finally {
    // 問題なければ完了報告を行う
    if (WorkStatus[0] == 'good') {
      const Text = `ステータス変更 ${SlackText} 問題なく完了しました`;
      await SlackPostMessage(Text);
    }
    // 問題があればエラー報告を行う
    if (WorkStatus[0] == 'error') {
      const Text = `ステータス変更 ${SlackText} エラーが発生しました\n@kushi_makoto 確認してください`;
      await SlackPostMessage(Text);
    }
    await RPA.WebBrowser.quit();
  }
}

Start();

async function Work() {
  const firstSheetData = [];
  const SheetWorkingRow = [];
  const LoopFlag = ['true'];
  const Text = `ステータス変更 ${SlackText} 開始します`;
  await SlackPostMessage(Text);
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

async function SlackPostMessage(Text) {
  await RPA.Logger.info('Slackに通知を送ったと想定' + Text);
  /*
  await RPA.Slack.chat.postMessage({
    channel: BotChannel,
    token: BotToken,
    text: Text,
    icon_emoji: ':snowman:',
    username: 'p1'
  });
  */
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

// 指定のページ(タブ)に移動する
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
  // 目的のタブをクリックする
  await TargetTabClick();
  await RPA.sleep(500);
  // たまにページが表示されないことがあるため、15秒待って出ない時はスキップする
  try {
    const CheckBox = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementsLocated({
        className: 'checkbox-cell'
      }),
      15000
    );
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
  for (let v = 2; v < 10; v++) {
    const Allbrake = ['false'];
    try {
      for (let NewNumber = 1; NewNumber < 101; NewNumber++) {
        var ID = await RPA.WebBrowser.findElementByCSSSelector(
          `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`
        );
        const IDText = await ID.getText();
        await RPA.Logger.info(IDText);
        if (IDText == SheetData[2]) {
          // 親ループもブレイクさせる
          Allbrake[0] = 'true';
          await RPA.Logger.info('ID一致しました');
          await RPA.WebBrowser.scrollTo({
            selector: `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td:nth-child(3)`
          });
          await RPA.Logger.info('一致したID の場所へスクロールしました');
          await RPA.sleep(200);
          // スイッチトグル を取得
          const SwitchToggle = await RPA.WebBrowser.findElementByCSSSelector(
            `#listTableAdGroup > tbody > tr:nth-child(${NewNumber}) > td.switch-box > switch-toggle > div`
          );
          // スイッチの状態 を取得
          const SwitchStatusList = await RPA.WebBrowser.findElementsByClassName(
            `switch ng-pristine ng-untouched ng-valid ng-not-empty`
          );
          const ListAttribute = await SwitchStatusList[
            Number(NewNumber) - 1
          ].getAttribute('class');
          await RPA.Logger.info(ListAttribute);
          // スイッチの状態が ON・OFF のどちらかなのか判定する
          if (ListAttribute.indexOf('active') > 0) {
            await RPA.Logger.info('現ステータス 【 ON 】 になっています');
            if (SheetData[3] == '無効') {
              await RPA.WebBrowser.mouseClick(SwitchToggle);
              await RPA.Logger.info('【 無効 】にしました');
            }
          } else {
            await RPA.Logger.info('現ステータス 【 OFF 】 になっています');
            if (SheetData[3] == '有効') {
              await RPA.WebBrowser.mouseClick(SwitchToggle);
              await RPA.Logger.info('【 有効 】にしました');
            }
          }
          await PasteSheet('完了', SheetWorkingRow);
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
