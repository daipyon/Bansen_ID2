import RPA from 'ts-rpa';
import { By, WebElement } from 'selenium-webdriver';

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【番宣 ID発行②（番宣）】発行完了しました`];

// スプレッドシートIDとシート名
// const mySSID = process.env.My_SheetID;
const SSID = process.env.Bansen_ID2_SheetID;
const SSName = process.env.Bansen_ID2_SheetName;
const SSName2 = process.env.Bansen_ID2_SheetName2;
// 作業するスプレッドシートから読み込む行数を記載
const StartRow = 3;
const LastRow = 200;

// 作業対象行とデータを取得
let SheetData;
let WorkData;
let Row;
// キャンペーンIDの有無を保持する変数
let CpnID;

// エラー発生時のテキストを格納
const Error_Text = [];
const errortext = 'エラー';

async function Start() {
  if (Error_Text.length == 0) {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    RPA.Logger.level = 'INFO';
    await RPA.Google.authorize({
      // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    SheetData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!A${StartRow}:AG${LastRow}`
    });
    // 作業用フラグ(AJ列)が"完了"の行まで回す
    for (let i in SheetData) {
      // CPN(AE列)を取得
      CpnID = SheetData[i][30];
      await RPA.Logger.info(`キャンペーンID → ${CpnID}`);
      Row = Number(i) + 3;
      if (SheetData[i][32] == errortext) {
        await RPA.Logger.info(
          `${Row} 行目のステータスが"エラー"ですのでスキップします`
        );
      } else if (SheetData[i][32] == '初期') {
        await RPA.Logger.info(`この行の作業を開始します → ${Row}`);
        // シートから作業対象行のデータを取得
        WorkData = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!A${Row}:AF${Row}`
        });
        await RPA.Logger.info(WorkData);
        // 作業用フラグ(AJ列)に"作業中"と記載
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!AJ${Row}:AJ${Row}`,
          values: [['作業中']]
        });
        await RPA.Logger.info(
          `${Row} 行目のステータスを"作業中"に変更しました`
        );
        await Work();
      } else if (SheetData[i][32] == '完了') {
        break;
      }
    }
  }
  // エラー発生時の処理
  if (Error_Text.length > 0) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(Error_Text);
    Slack_Text[0] = `【番宣 ID発行②（番宣）】でエラー発生しました\n${Error_Text}`;
    await RPA.WebBrowser.takeScreenshot();
  }
  await RPA.Logger.info(Slack_Text[0]);
  await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

let FirstLoginFlag = 'true';
async function Work() {
  try {
    // 配信先が「リニア」、配信方式が「埋め配信」
    const pattern1 = WorkData[0][5] == 'リニア' && WorkData[0][6] == '埋め配信';
    // 配信先が「リニア」、配信方式が「KPI配信」
    const pattern2 = WorkData[0][5] == 'リニア' && WorkData[0][6] == 'KPI配信';
    // 配信先が「ビデオ,タイムシフト」、配信方式が「KPI配信」
    const pattern3 =
      WorkData[0][5] == 'ビデオ,タイムシフト' && WorkData[0][6] == 'KPI配信';
    // 配信先が「ビデオ,タイムシフト」、配信方式が「指定配信」
    const pattern4 =
      WorkData[0][5] == 'ビデオ,タイムシフト' && WorkData[0][6] == '指定配信';
    // 配信先が「リニア」、配信方式が「指定配信」
    const pattern5 = WorkData[0][5] == 'リニア' && WorkData[0][6] == '指定配信';
    // AAAMSにログイン
    if (FirstLoginFlag == 'true') {
      await AAAMSLogin();
    }
    if (FirstLoginFlag == 'false') {
      await AAAMSLogin2();
    }
    // 一度ログインしたら、以降はログインページをスキップ
    FirstLoginFlag = 'false';
    // キャンペーンIDが発番されていない場合、キャンペーン作成からスタート
    if (CpnID == ``) {
      if (pattern5) {
        await RPA.Logger.info(`パターン5 のためスキップします`);
        const ErrorText = [
          [
            errortext,
            `配信先が「リニア」、配信方式が「指定配信」のためスキップしました`
          ]
        ];
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!AJ${Row}:AK${Row}`,
          values: ErrorText
        });
        await Start();
      }
      // キャンペーン作成を開始
      await CampaignStart(pattern1, pattern2, pattern3, pattern4);
      // 共通項目のエラー判定
      await CommonJudgeError();
      // 共通項目②を入力
      await Common2Input();
      // 共通項目②のエラー判定
      await Common2JudgeError();
      // パターン別に入力
      await PatternInput(pattern1, pattern2, pattern3, pattern4);
      // クラスタ(M列)を選択
      await Cluster();
      // パターン別のエラー判定
      await PatternJudgeError(pattern1, pattern2, pattern3, pattern4);
      // クラスタのエラー判定
      await ClusterJudgeError();
      // キャンペーンIDを発番し、取得
      await GetCampaignId();
      // 広告作成を開始
      await AdvertisementStart();
      // 広告作成のエラー判定
      await AdJudgeError();
      // 広告IDを発番
      await GetAdvertisementId();
      // 既にキャンペーンIDが発番されている場合、広告作成からスタート
    } else {
      // キャンペーンIDで検索
      await CampaignIdSerach();
      await AdvertisementStart();
      await AdJudgeError();
      await GetAdvertisementId();
    }
  } catch (error) {
    Error_Text[0] = error;
    await Start();
  }
}

async function SlackPost(Text) {
  await RPA.Slack.chat.postMessage({
    token: Slack_Token,
    channel: Slack_Channel,
    text: `${Text}`
  });
}

async function AAAMSLogin() {
  await RPA.Logger.info('AAAMSにログインします');
  // 本番用
  await RPA.WebBrowser.get(process.env.AAAMS_Login_URL);

  // テスト用
  // await RPA.WebBrowser.get(process.env.STG_AAAMS_Login_URL);
  await RPA.sleep(2000);
  const AAAMS_loginID_ele = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      name: 'email'
    }),
    15000
  );
  await RPA.WebBrowser.sendKeys(AAAMS_loginID_ele, [
    // 本番用
    `${process.env.Bansen_ID2_AAAMSID}`

    // テスト用
    // `${process.env.Bansen_ID2_STG_AAAMSID}`
  ]);
  const AAAMS_loginPW_ele = await RPA.WebBrowser.driver.findElement(
    By.name('password')
  );
  await RPA.WebBrowser.sendKeys(AAAMS_loginPW_ele, [
    // 本番用
    `${process.env.Bansen_ID2_AAAMSPW}`

    // テスト用
    // `${process.env.Bansen_ID2_STG_AAAMSPW}`
  ]);
  const AAAMS_LoginNextButton = await RPA.WebBrowser.findElementByClassName(
    'auth0-lock-submit'
  );
  await RPA.WebBrowser.mouseClick(AAAMS_LoginNextButton);
  while (0 == 0) {
    await RPA.sleep(5000);
    const AutoReload = await RPA.WebBrowser.findElementsByClassName(
      'Modal__message___bjHRV'
    );
    if (AutoReload.length > 0) {
      await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
      break;
    }
  }
  await UpdateScreenThrough();
  await RPA.sleep(3000);
  const NextButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className:
        'Button__button___2uDT- Button__ok___1qMMy Modal__buttonOkay___2ewPw'
    }),
    15000
  );
  await RPA.WebBrowser.mouseClick(NextButton);
  await RPA.sleep(3000);
  await RPA.Logger.info('番宣アカウントを直接呼び出します');
  // 本番用
  await RPA.WebBrowser.get(process.env.AAAMS_Account_3);

  // テスト用
  // await RPA.WebBrowser.get(process.env.STG_AAAMS_Account);
  await RPA.sleep(3000);
}

// 更新画面をスルーする関数
async function UpdateScreenThrough() {
  // チャンネル更新画面が出た場合
  try {
    const ChannelAlart = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'Modal__title___PGAs6'
      }),
      15000
    );
    const ChanneAlartText = await ChannelAlart.getText();
    if (ChanneAlartText == '下記更新されました。設定を確認してください。') {
      await RPA.Logger.info('更新画面が出ました');
      const ChannelAlartButton: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('Button__button___2uDT- Button__ok___1qMMy Modal__buttonOkay___2ewPw')[1]`
      );
      await RPA.WebBrowser.mouseClick(ChannelAlartButton);
      await RPA.sleep(2000);
    }
  } catch {
    await RPA.Logger.info('更新画面は出ませんでした');
  }
}

async function AAAMSLogin2() {
  await RPA.Logger.info('番宣アカウントを直接呼び出します');
  await RPA.WebBrowser.get(process.env.AAAMS_Account_3);
  await RPA.sleep(3000);
}

async function CampaignStart(pattern1, pattern2, pattern3, pattern4) {
  await RPA.Logger.info('キャンペーン 作成します');
  // 右上のキャンペーン作成を押下
  const CampaignCreateButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('ContentsGroup__actions___2Bw_9')[0].children[0]`
  );
  await RPA.WebBrowser.mouseClick(CampaignCreateButton);
  await RPA.sleep(2000);
  await RPA.Logger.info('共通項目 を入力します');
  // キャンペーン名(A列)を入力
  await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      name: 'campaign.name'
    }),
    15000
  );
  const CampaignName: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.name')[0].children[1].children[0].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CampaignName, [WorkData[0][0]]);
  // 有効期間をクリック
  const DateRange: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.dateRange')[0].children[1].children[0].children[0].children[0].children[0]`
  );
  await RPA.WebBrowser.mouseClick(DateRange);
  // 有効期間：開始(B列)を入力
  const DateRangeStart: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('daterangepicker_start')[1]`
  );
  await DateRangeStart.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeStart, [WorkData[0][1]]);
  await RPA.sleep(300);
  // 有効期間：終了(C列)を入力
  const DateRangeEnd: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('daterangepicker_end')[1]`
  );
  await DateRangeEnd.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeEnd, [WorkData[0][2]]);
  await RPA.sleep(300);
  // 有効期間のOKボタンを押下
  const OKButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('applyBtn btn btn-sm btn-success')[1]`
  );
  await RPA.WebBrowser.mouseClick(OKButton);
  await RPA.sleep(700);
  // 予約放送枠ID(D列)を入力
  const CampaignSlotId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.slotIds[0]')[0].children[0].children[0]`
  );
  if (WorkData[0][3] == 'なし') {
    await RPA.Logger.info('予約放送枠ID が "なし" のためスルーします');
  } else {
    await RPA.WebBrowser.sendKeys(CampaignSlotId, [WorkData[0][3]]);
  }
  // キャンペーン種別(E列)を入力
  const CampaignPromo: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.promo')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CampaignPromo, [WorkData[0][4]]);
  await RPA.WebBrowser.sendKeys(CampaignPromo, [RPA.WebBrowser.Key.ENTER]);
  // 配信先設定(F列)を選択
  await CampaignDeliver(pattern1, pattern2, pattern3, pattern4);
  // 配信方式設定(G列)を入力
  const CampaignPlacement: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.placement')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CampaignPlacement, [WorkData[0][6]]);
  await RPA.WebBrowser.sendKeys(CampaignPlacement, [RPA.WebBrowser.Key.ENTER]);
}

async function CampaignDeliver(pattern1, pattern2, pattern3, pattern4) {
  await RPA.Logger.info('配信先設定 を選択します');
  const Linear: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.deliver')[0].children[1].children[0].children[1].children[0]`
  );
  // 配信先設定の指定がない場合
  if (WorkData[0][5].length < 1) {
    await RPA.WebBrowser.mouseClick(Linear);
  } else {
    if (pattern1 || pattern2) {
      await RPA.Logger.info('デフォルトの状態のためスルーします');
    }
    if (pattern3 || pattern4) {
      await RPA.Logger.info('ビデオ・タイムシフトにチェックします');
      await RPA.WebBrowser.mouseClick(Linear);
      const Video: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.deliver')[0].children[1].children[0].children[1].children[1]`
      );
      await RPA.WebBrowser.mouseClick(Video);
      const TimeShift: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.deliver')[0].children[1].children[0].children[1].children[2]`
      );
      await RPA.WebBrowser.mouseClick(TimeShift);
    }
  }
}

const ErrorText1 = '同じキャンペーン名が既に存在しています';
const ErrorText2 = 'キャンペーン名が65文字以上です';
const ErrorText3 = '必須入力です';
const ErrorText4 = '必須入力です';
const ErrorText5 = 'シリーズIDが存在しません';

async function CommonJudgeError() {
  await RPA.Logger.info('共通項目 エラー判定開始...');
  const DoubleCampaignName: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.name')[0].children[0]`
  );
  const DoubleCampaignNameText = await DoubleCampaignName.getText();
  if (
    DoubleCampaignNameText ==
    'キャンペーン名必須\n同じキャンペーン名が既に存在しています'
  ) {
    await RPA.Logger.info(
      `エラー【キャンペーン名：同じキャンペーン名が既に存在しています】`
    );
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AL${Row}:AL${Row}`,
      values: [[ErrorText1]]
    });
  }
  if (
    DoubleCampaignNameText == 'キャンペーン名必須\n最大100文字まで入力できます'
  ) {
    await RPA.Logger.info(
      `エラー【キャンペーン名：最大100文字まで入力できます】`
    );
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AL${Row}:AL${Row}`,
      values: [[ErrorText2]]
    });
  }
  if (DoubleCampaignNameText == 'キャンペーン名必須\n必須入力です') {
    await RPA.Logger.info(`エラー【キャンペーン名：${ErrorText4}】`);
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AL${Row}:AL${Row}`,
      values: [[ErrorText3]]
    });
  }
  if (
    WorkData[0][1].length < 1 ||
    WorkData[0][2].length < 1 ||
    (WorkData[0][1].length < 1 && WorkData[0][2].length < 1)
  ) {
    await RPA.Logger.info(`エラー【有効期間：必須項目です】`);
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AM${Row}:AM${Row}`,
      values: [['必須項目です']]
    });
  }
  // 1秒以上スリープを入れないと予約放送枠IDの判定ができない
  await RPA.sleep(1000);
  const NotCampaignSlotId = await RPA.WebBrowser.findElementByClassName(
    'FieldSlotIds__slot___1dJWu'
  );
  const NotCampaignSlotIdText = await NotCampaignSlotId.getText();
  if (NotCampaignSlotIdText == '該当する放送枠が存在しません') {
    // || NotCampaignSlotIdText == '既に放送終了している放送枠です'
    await RPA.Logger.info(`エラー【予約放送枠ID：${NotCampaignSlotIdText}】`);
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AN${Row}:AN${Row}`,
      values: [[NotCampaignSlotIdText]]
    });
  }
  if (WorkData[0][5].length < 1) {
    await RPA.Logger.info(`エラー【配信先設定：${ErrorText4}】`);
    await ErrorSetValue(`AJ${Row}`);
    await RequiredInputSetValue(`AO${Row}`);
  }
  await SheetJudgeError(Row);
}

async function Common2Input() {
  await RPA.Logger.info('共通のデータ を入力します');
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.dayHourTempId')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId, [WorkData[0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId, [
    RPA.WebBrowser.Key.ENTER
  ]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Field__value___3-hi8 Field__flex___2MAX7')[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId, [WorkData[0][8]]);
  await RPA.sleep(1000);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName(
    'Select-option'
  );
  const FilterTempValueText = await Promise.all(
    FirstFilterTempValue.map(async elm => await elm.getText())
  );
  await RPA.Logger.info(
    '配信先フィルタテンプレート一覧 → ' + FilterTempValueText
  );
  for (let i in FilterTempValueText) {
    if (WorkData[0][8] == FilterTempValueText[i]) {
      await RPA.Logger.info(`${FilterTempValueText[i]} → 一致しました`);
      // const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(
      //   `/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/div/div[${Number(
      //     i
      //   ) + Number(1)}]`
      // );
      await RPA.sleep(1000);
      // await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      await RPA.WebBrowser.sendKeys(CampaignFilterTempId, [
        RPA.WebBrowser.Key.ENTER
      ]);
      FilterTempFlag[0] = false;
      break;
    }
    await RPA.WebBrowser.sendKeys(CampaignFilterTempId, [
      RPA.WebBrowser.Key.DOWN
    ]);
  }
  // 取得した値を分割
  const ValuesJ = WorkData[0][9].split(',');
  await RPA.Logger.info(ValuesJ);
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.attributeIds')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
  );
  for (let i in ValuesJ) {
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds, [ValuesJ[i]]);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds, [
      RPA.WebBrowser.Key.ENTER
    ]);
  }
}

async function Common2JudgeError() {
  if (WorkData[0][6].length < 1) {
    await RPA.Logger.info(`エラー【配信方式設定：${ErrorText4}】`);
    await ErrorSetValue(`AJ${Row}`);
    await RequiredInputSetValue(`AP${Row}`);
  } else {
    try {
      const SelectCampaignPlacement: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.placement')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[0]`
      );
      const SelectCampaignPlacementText = await SelectCampaignPlacement.getText();
      if (SelectCampaignPlacementText == '配信方式設定') {
        await RPA.Logger.info(
          `エラー【配信方式設定：配信方式が選択されていません】`
        );
        await ErrorSetValue(`AJ${Row}`);
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!AP${Row}:AP${Row}`,
          values: [['配信方式が選択されていません']]
        });
      }
    } catch {}
  }
  if (WorkData[0][7].length < 1) {
    await RPA.Logger.info(
      `エラー【配信先曜日×時間帯テンプレート：${ErrorText4}】`
    );
    await ErrorSetValue(`AJ${Row}`);
    await RequiredInputSetValue(`AQ${Row}`);
  } else {
    try {
      const SelectCampaignDayHourTempId: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.dayHourTempId')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[0]`
      );
      const SelectCampaignDayHourTempIdText = await SelectCampaignDayHourTempId.getText();
      if (
        SelectCampaignDayHourTempIdText == '配信先曜日×時間帯テンプレートを選択'
      ) {
        await RPA.Logger.info(
          `エラー【配信先曜日×時間帯テンプレート：テンプレートが選択されていません】`
        );
        await ErrorSetValue(`AJ${Row}`);
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!AQ${Row}:AQ${Row}`,
          values: [['テンプレートが選択されていません']]
        });
      }
    } catch {}
  }
  if (WorkData[0][8].length < 1) {
    await RPA.Logger.info(
      `エラー【配信先フィルタテンプレート：${ErrorText4}】`
    );
    await ErrorSetValue(`AJ${Row}`);
    await RequiredInputSetValue(`AR${Row}`);
  } else {
    try {
      const SelectCampaignFilterTempId: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('Field__value___3-hi8 Field__flex___2MAX7')[1].children[0].children[0].children[0].children[0].children[0].children[0].children[0]`
      );
      const SelectCampaignFilterTempIdText = await SelectCampaignFilterTempId.getText();
      if (
        SelectCampaignFilterTempIdText ==
          'リニア配信先フィルタテンプレートを選択' ||
        SelectCampaignFilterTempIdText ==
          'ビデオ配信先フィルタテンプレートを選択'
      ) {
        await RPA.Logger.info(
          `エラー【配信先フィルタテンプレート：テンプレートが選択されていません】`
        );
        await ErrorSetValue(`AJ${Row}`);
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!AR${Row}:AR${Row}`,
          values: [['テンプレートが選択されていません']]
        });
      }
    } catch {}
  }
  await SheetJudgeError(Row);
}

async function Level(WorkData) {
  // 隣接許容レベル・広告(K列)を入力
  if (WorkData[0][10].length < 1) {
    const CroseButton: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.adjacentAcceptable.ad')[0].children[1].children[0].children[0].children[0].children[0].children[1].children[1]`
    );
    await RPA.WebBrowser.mouseClick(CroseButton);
  } else {
    const CampaignAdjacentAcceptableAd: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.adjacentAcceptable.ad')[0].children[1].children[0].children[0].children[0].children[0].children[1].children[0].children[1].children[0]`
    );
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd, [
      WorkData[0][10]
    ]);
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd, [
      RPA.WebBrowser.Key.ENTER
    ]);
  }
  // 隣接許容レベル・クリエイティブ(L列)を選択
  if (WorkData[0][11].length < 1) {
    const CroseButton2: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.adjacentAcceptable.creative')[0].children[1].children[0].children[0].children[0].children[0].children[1].children[1]`
    );
    await RPA.WebBrowser.mouseClick(CroseButton2);
  } else {
    const CampaignAdjacentAcceptableCreative: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.adjacentAcceptable.creative')[0].children[1].children[0].children[0].children[0].children[0].children[1].children[0].children[1].children[0]`
    );
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative, [
      WorkData[0][11]
    ]);
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative, [
      RPA.WebBrowser.Key.ENTER
    ]);
  }
}

async function LevelJudgeError(Row) {
  try {
    const NotCampaignAdjacentAcceptableAd: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.adjacentAcceptable.ad')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[0]`
    );
    const NotCampaignAdjacentAcceptableAdText = await NotCampaignAdjacentAcceptableAd.getText();
    if (
      NotCampaignAdjacentAcceptableAdText ==
      '広告に対しての隣接許容レベルを選択'
    ) {
      await RPA.Logger.info(`エラー【隣接許容レベル・広告：${ErrorText4}】`);
      await ErrorSetValue(`AJ${Row}`);
      await RequiredInputSetValue(`AS${Row}`);
    }
  } catch {}
  try {
    const NotCampaignAdjacentAcceptableCreative: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.adjacentAcceptable.creative')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[0]`
    );
    const NotCampaignAdjacentAcceptableCreativeText = await NotCampaignAdjacentAcceptableCreative.getText();
    if (
      NotCampaignAdjacentAcceptableCreativeText ==
      'クリエイティブに対しての隣接許容レベルを選択'
    ) {
      await RPA.Logger.info(
        `エラー【隣接許容レベル・クリエイティブ：${ErrorText4}}】`
      );
      await ErrorSetValue(`AJ${Row}`);
      await RequiredInputSetValue(`AT${Row}`);
    }
  } catch {}
}

async function PatternInput(pattern1, pattern2, pattern3, pattern4) {
  if (pattern1 || pattern2) {
    // レベルを入力
    await Level(WorkData);
    // レベルのエラー判定
    await LevelJudgeError(Row);
  }
  if (pattern1 || pattern2 || pattern3) {
    // 指定シリーズNG(R列)を入力
    const CampaignSeriesNg = await RPA.WebBrowser.findElementByClassName(
      'FieldText__textarea___3vdMx'
    );
    await RPA.WebBrowser.sendKeys(CampaignSeriesNg, [WorkData[0][17]]);
    await RPA.WebBrowser.sendKeys(CampaignSeriesNg, [RPA.WebBrowser.Key.ENTER]);
  }
  if (pattern1) {
    await RPA.Logger.info('パターン 1 です');
    // 比率(N列)を入力
    const CampaignStockWeight: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.stockWeight')[0].children[1].children[0].children[0]`
    );
    await RPA.WebBrowser.sendKeys(CampaignStockWeight, [WorkData[0][13]]);
    await RPA.WebBrowser.sendKeys(CampaignStockWeight, [
      RPA.WebBrowser.Key.ENTER
    ]);
  }
  if (pattern2 || pattern3 || pattern4) {
    // 総目標imp(P列)を入力
    const GoalGimp: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('goal.gimp')[0].children[1].children[0].children[0]`
    );
    await RPA.WebBrowser.sendKeys(GoalGimp, [WorkData[0][15]]);
    await RPA.WebBrowser.sendKeys(GoalGimp, [RPA.WebBrowser.Key.ENTER]);
    //「FQキャップ」に「3」と入力（デフォルト）
    const FqCap: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('FieldText__inputWrap___24SZl FieldText__unitWrap___2SOu6')[2].children[0]`
    );
    await FqCap.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(FqCap, [`3`]);
    await RPA.sleep(300);
  }
  if (pattern2) {
    await RPA.Logger.info('パターン 2 です');
    // 優先度(O列)を入力
    const GoalPriority: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('goal.priority')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
    );
    await RPA.WebBrowser.sendKeys(GoalPriority, [WorkData[0][14]]);
    await RPA.WebBrowser.sendKeys(GoalPriority, [RPA.WebBrowser.Key.ENTER]);
    // カスタムオーディエンス(Q列)を入力
    const CustomAudience: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.customAudienceString.id')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
    );
    await RPA.WebBrowser.sendKeys(CustomAudience, [WorkData[0][16]]);
    await RPA.WebBrowser.sendKeys(CustomAudience, [RPA.WebBrowser.Key.ENTER]);
  }
  if (pattern3 || pattern4) {
    await RPA.sleep(1000);
    // 訴求シリーズID(S列)を入力
    const CampaignPromoSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.promoSeriesId[0]')[0].children[0].children[0]`
    );
    await RPA.WebBrowser.sendKeys(CampaignPromoSeriesId, [WorkData[0][18]]);
  }
  if (pattern3) {
    await RPA.Logger.info('パターン 3 です');
  }
  if (pattern4) {
    await RPA.Logger.info('パターン 4 です');
    // 指定シリーズID(T列)を入力
    const CampaignSpecifiedSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('FieldText__textarea___3vdMx')[1]`
    );
    const Text1 = await encodeURI(WorkData[0][19]);
    await RPA.WebBrowser.driver.executeScript(
      // `document.querySelectorAll("#reactroot > div > div.Modal__overlay___3CdlJ.Modal__full___3Mi9N.Modal__opened___-zyiB > div.Modal__modal___15eHS.Modal__full___3Mi9N > div.Modal__body___2lTdW.Modal__full___3Mi9N > div > form > div > div:nth-child(4) > div.FieldGroup__container___3fkr0 > div > div:nth-child(10) > div.Field__value___3-hi8 > div > div.Field__container___2rw9q.Field__column___2-TfA.Field__labelInline___2GdCv.Field__noMargin___1oAyZ.Field__alignNone___1dlxQ > div.Field__value___3-hi8 > textarea")[0].value = decodeURI("${Text1}");`
      `document.getElementsByClassName('FieldText__textarea___3vdMx')[1].value = decodeURI("${Text1}");`
    );
    await RPA.WebBrowser.sendKeys(CampaignSpecifiedSeriesId, [
      RPA.WebBrowser.Key.ENTER
    ]);
  }
}

async function Cluster() {
  await RPA.Logger.info('クラスタ を選択します');
  await RPA.Logger.info(WorkData[0][12]);
  if (WorkData[0][12].includes('全て選択') == true) {
    await RPA.Logger.info(' "全て選択" のためスルーします');
  } else {
    await RPA.Logger.info('記載されているものにチェックを入れます');
    // "全て選択"をクリックし、全てのチェックを外す
    const All: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.clusterIds')[0].children[1].children[0].children[0]`
    );
    await RPA.WebBrowser.mouseClick(All);
    // 取得した値を分割
    const ValuesM = WorkData[0][12].split(',');
    if (ValuesM.includes('F1') == true) {
      await RPA.Logger.info(' "F1" にチェックを入れます');
      // "F1"をクリック
      const F1: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.clusterIds')[0].children[1].children[0].children[1].children[0]`
      );
      await RPA.WebBrowser.mouseClick(F1);
    }
    if (ValuesM.includes('F2') == true) {
      await RPA.Logger.info(' "F2" にチェックを入れます');
      // "F2"をクリック
      const F2: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.clusterIds')[0].children[1].children[0].children[1].children[1]`
      );
      await RPA.WebBrowser.mouseClick(F2);
    }
    if (ValuesM.includes('M1') == true) {
      await RPA.Logger.info(' "M1" にチェックを入れます');
      // "M1"をクリック
      const M1: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.clusterIds')[0].children[1].children[0].children[1].children[2]`
      );
      await RPA.WebBrowser.mouseClick(M1);
    }
    if (ValuesM.includes('M2') == true) {
      await RPA.Logger.info(' "M2" にチェックを入れます');
      // "M2"をクリック
      const M2: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.clusterIds')[0].children[1].children[0].children[1].children[3]`
      );
      await RPA.WebBrowser.mouseClick(M2);
    }
    if (ValuesM.includes('teen') == true) {
      await RPA.Logger.info(' "teen" にチェックを入れます');
      // "teen"をクリック
      const Teen: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.clusterIds')[0].children[1].children[0].children[1].children[4]`
      );
      await RPA.WebBrowser.mouseClick(Teen);
    }
    if (ValuesM.includes('other') == true) {
      await RPA.Logger.info(' "other" にチェックを入れます');
      // "other"をクリック
      const Other: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.clusterIds')[0].children[1].children[0].children[1].children[5]`
      );
      await RPA.WebBrowser.mouseClick(Other);
    }
  }
}

async function PatternJudgeError(pattern1, pattern2, pattern3, pattern4) {
  if (pattern1 || pattern2 || pattern3) {
    if (WorkData[0][17].length < 1) {
      await RPA.Logger.info('指定シリーズNG の記載がないためスルーします');
    } else {
      const NotCampaignSeriesNg: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.seriesIds[0]')[0].children[1]`
      );
      const NotCampaignSeriesNgText = await NotCampaignSeriesNg.getText();
      if (NotCampaignSeriesNgText == ErrorText5) {
        await RPA.Logger.info(`エラー【指定シリーズNG：${ErrorText5}】`);
        await ErrorSetValue(`AJ${Row}`);
        await NotSeriesId(`AU${Row}`);
      }
    }
  }
  if (pattern1) {
    await RPA.Logger.info('パターン 1 エラー判定開始...');
    const NotCampaignStockWeight: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.stockWeight')[0].children[0]`
    );
    const NotCampaignStockWeightText = await NotCampaignStockWeight.getText();
    if (NotCampaignStockWeightText == '比率必須必須入力です') {
      await RPA.Logger.info(`エラー【比率：${ErrorText4}】`);
      await ErrorSetValue(`AJ${Row}`);
      await RequiredInputSetValue(`AW${Row}`);
    }
  }
  if (pattern2 || pattern3 || pattern4) {
    const NotGoalGimp: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('goal.gimp')[0].children[0]`
    );
    const NotGoalGimpText = await NotGoalGimp.getText();
    if (NotGoalGimpText == '総目標imp必須必須入力です') {
      await RPA.Logger.info(`エラー【総目標imp：${ErrorText4}】`);
      await ErrorSetValue(`AJ${Row}`);
      await RequiredInputSetValue(`AY${Row}`);
    }
  }
  if (pattern2) {
    await RPA.Logger.info('パターン 2 エラー判定開始...');
    try {
      const NotGoalPriority: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('goal.priority')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[0]`
      );
      const NotGoalPriorityText = await NotGoalPriority.getText();
      if (NotGoalPriorityText == '優先度') {
        await RPA.Logger.info(`エラー【優先度：${ErrorText4}】`);
        await ErrorSetValue(`AJ${Row}`);
        await RequiredInputSetValue(`AX${Row}`);
      }
    } catch {}
  }
  if (pattern3 || pattern4) {
    const NotCampaignPromoSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.promoSeriesId[0]')[0].children[1]`
    );
    const NotCampaignPromoSeriesIdText = await NotCampaignPromoSeriesId.getText();
    if (NotCampaignPromoSeriesIdText == ErrorText5) {
      await RPA.Logger.info(`エラー【訴求シリーズID：${ErrorText5}】`);
      await ErrorSetValue(`AJ${Row}`);
      await NotSeriesId(`BA${Row}`);
    }
  }
  if (pattern3) {
    await RPA.Logger.info('パターン 3 エラー判定開始...');
  }
  if (pattern4) {
    await RPA.Logger.info('パターン 4 エラー判定開始...');
    if (WorkData[0][19].length < 1) {
      await RPA.Logger.info('エラー【指定シリーズID】IDの記載がありません');
      await ErrorSetValue(`AJ${Row}`);
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName}!AZ${Row}:AZ${Row}`,
        values: [['IDの記載がありません']]
      });
    } else {
      const NotCampaignSpecifiedSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('campaign.vodSeriesIds[0]')[0].children[1]`
      );
      const NotCampaignSpecifiedSeriesIdText = await NotCampaignSpecifiedSeriesId.getText();
      if (NotCampaignSpecifiedSeriesIdText == ErrorText5) {
        await RPA.Logger.info(`エラー【指定シリーズID：${ErrorText5}】`);
        await ErrorSetValue(`AJ${Row}`);
        await NotSeriesId(`AZ${Row}`);
      }
    }
  }
}

async function ClusterJudgeError() {
  if (WorkData[0][12].length < 1) {
    await RPA.Logger.info(`エラー【クラスタ：${ErrorText4}】`);
    await ErrorSetValue(`AJ${Row}`);
    await RequiredInputSetValue(`AV${Row}`);
  }
  await SheetJudgeError(Row);
}

async function GetCampaignId() {
  await RPA.Logger.info('キャンペーン ID発番します');
  const OKButton2: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Modal__footer___aDbwS Modal__full___3Mi9N')[0].children[1]`
  );
  await RPA.WebBrowser.mouseClick(OKButton2);
  await RPA.Logger.info('発番中...');
  await RPA.sleep(5000);
  await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: `Table__cell___3h99d Table__campaignName___z9GnE`
    }),
    15000
  );
  for (let i = 0; i <= 14; i++) {
    const CampaignName2: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('Table__cell___3h99d Table__campaignName___z9GnE')[${i}]`
    );
    const CampaignName2Text = await CampaignName2.getText();
    await RPA.Logger.info(
      `取得したキャンペーン名　　　　　　　　　→　${CampaignName2Text}`
    );
    if (CampaignName2Text == WorkData[0][0]) {
      await RPA.WebBrowser.mouseClick(CampaignName2);
      break;
    }
  }
  await RPA.Logger.info(
    `現在保持しているデータのキャンペーン名　→　${WorkData[0][0]}`
  );
  await RPA.Logger.info('キャンペーン名が一致しましたので、IDを取得します');
  await RPA.sleep(5000);
  const CampaignId = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'CampaignDetailData__item___3MVsT'
    }),
    15000
  );
  const CampaignIdText = await CampaignId.getText();
  await RPA.Logger.info(CampaignIdText);
  // キャンペーン・広告貼り付けシートのBE列に発番したキャンペーンIDを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!BE${Row - 1}:BE${Row - 1}`,
    values: [[CampaignIdText]]
  });
  // キャンペーン名が同じものがあればその行にもキャンペーンIDを記載
  const JudgeCampaignName = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!X2:X3000`
  });
  for (let i in JudgeCampaignName) {
    if (JudgeCampaignName[i][0] == WorkData[0][0]) {
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName2}!BE${Number(i) + 2}:BE${Number(i) + 2}`,
        values: [[CampaignIdText]]
      });
    }
  }
  await RPA.Logger.info('キャンペーン 作成完了しました');
  // 右下の広告作成を押下
  const AdvertisementCreateButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('ContentsGroup__childHeader___bAbEP')[0].children[1].children[0]`
  );
  await RPA.WebBrowser.mouseClick(AdvertisementCreateButton);
  await RPA.sleep(3000);
}

let LoadingFlag = 'false';
async function CampaignIdSerach() {
  while (LoadingFlag == 'false') {
    const JudgeCampaignId = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName2}!BE${Row - 1}:BE${Row - 1}`
    });
    // キャンペーンIDで検索
    await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        name: `campaign.ids`
      }),
      15000
    );
    const SearchByCampaignId: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('FieldText__input___2AKz1 FieldText__search___3XxRO')[1]`
    );
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(SearchByCampaignId, JudgeCampaignId[0]);
    await RPA.sleep(5000);
    try {
      const Loading = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          className: `Spinner__message___3UgRY`
        }),
        15000
      );
      const LoadingText = await Loading.getText();
      if (LoadingText == '読み込み中...') {
        await RPA.Logger.info(
          '検索結果が出ないため、ブラウザを更新し再検索します'
        );
        await RPA.WebBrowser.refresh();
      }
      await RPA.Logger.info('番宣アカウントを直接呼び出します');
      await RPA.WebBrowser.get(process.env.AAAMS_Account_3);
      await RPA.sleep(3000);
      // 代理店の箇所に「番宣」の文字が入っていない場合は再更新
      const AgencyName: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByName('agency')[0].children[1].children[0].children[0].children[0].children[1].children[0].children[0].children[0]`
      );
      const AgencyNameText = await AgencyName.getText();
      if (AgencyNameText == '代理店名で検索') {
        await RPA.WebBrowser.refresh();
      }
      await UpdateScreenThrough();
      await RPA.sleep(3000);
      const CampaignId2 = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          className: 'Table__cell___3h99d Table__campaignId___319zD'
        }),
        15000
      );
      const CampaignId2Text = await CampaignId2.getText();
      if (CampaignId2Text == JudgeCampaignId[0][0]) {
        LoadingFlag = 'true';
        await RPA.Logger.info(
          `取得したキャンペーンID　　      　 →　${CampaignId2Text}`
        );
        await RPA.Logger.info(
          `貼り付けシートのキャンペーンID　   →　${JudgeCampaignId[0][0]}`
        );
        await RPA.Logger.info(
          'キャンペーンIDが一致しましたので、広告作成を開始します'
        );
      }
    } catch {
      break;
    }
  }
  LoadingFlag = 'false';
  const CampaignName3 = await RPA.WebBrowser.findElementByClassName(
    'Table__cell___3h99d Table__campaignName___z9GnE'
  );
  await RPA.WebBrowser.mouseClick(CampaignName3);
  await RPA.sleep(3000);
  // 右下の広告作成を押下
  const AdvertisementCreateButton2 = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className:
        'Button__button___2uDT- Button__create___1F14z Button__production___v52Gd'
    }),
    15000
  );
  await RPA.WebBrowser.mouseClick(AdvertisementCreateButton2);
  await RPA.sleep(3000);
}

async function AdvertisementStart() {
  await RPA.Logger.info('広告 作成します');
  // クリエイティブID(X列)を入力
  const CreativeId = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'FieldText__input___2AKz1'
    }),
    15000
  );
  // クリエイティブIDがない場合、作業をスキップしてスタートに戻る
  if (WorkData[0][23].length < 1) {
    await RPA.Logger.info(
      `エラー【シートにIDの記載がありません】 作業をスキップします`
    );
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!BB${Row}:BB${Row}`,
      values: [['シートにIDの記載がありません']]
    });
    await Start();
  } else {
    await RPA.WebBrowser.sendKeys(CreativeId, [WorkData[0][23]]);
    await RPA.sleep(3000);
    // エラーが起きた場合、作業をスキップしてスタートに戻る
    const ApplyFlag = ['false'];
    try {
      const NotCreativeId = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          className: 'Table__message___2T1_9'
        }),
        15000
      );
      const NotCreativeIdText = await NotCreativeId.getText();
      if (NotCreativeIdText == '該当する項目が存在しません') {
        await RPA.Logger.info(
          `エラー【${NotCreativeIdText}】 作業をスキップします`
        );
        await ErrorSetValue(`AJ${Row}`);
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!BB${Row}:BB${Row}`,
          values: [[NotCreativeIdText]]
        });
        await Start();
      }
    } catch {
      ApplyFlag[0] = 'true';
      await RPA.Logger.info('次の処理に進みます');
    }
    const SelectCreative = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className:
          'Button__button___2uDT- Button__inlist___2iwq1 Button__production___v52Gd'
      }),
      15000
    );
    await RPA.WebBrowser.mouseClick(SelectCreative);
    // 配信期間をクリック
    const DateRange2: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('ads[0].dateRange')[0].children[1].children[0].children[0].children[0].children[0]`
    );
    await RPA.WebBrowser.mouseClick(DateRange2);
    // 配信期間：開始(AC列)を入力
    const DateRangeStart2 = await RPA.WebBrowser.driver.findElement(
      By.name('daterangepicker_start')
    );
    await DateRangeStart2.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeStart2, [WorkData[0][28]]);
    await RPA.sleep(300);
    // 配信期間：終了(AD列)を入力
    const DateRangeEnd2 = await RPA.WebBrowser.driver.findElement(
      By.name('daterangepicker_end')
    );
    await DateRangeEnd2.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeEnd2, [WorkData[0][29]]);
    await RPA.sleep(300);
    // 配信期間のOKボタンを押下
    const OKButton3 = await RPA.WebBrowser.findElementByClassName(
      'applyBtn btn btn-sm btn-success'
    );
    await RPA.WebBrowser.mouseClick(OKButton3);
    await RPA.sleep(1000);
    // 一度入力されている開始・終了期間をそれぞれ取得
    const DateRange2Text = await DateRange2.getText();
    await RPA.Logger.info(`AAAMSの開始日  → ${DateRange2Text}`);
    await RPA.Logger.info(`シートの開始日 → ${WorkData[0][28]}`);
    const DateRange3: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('ads[0].dateRange')[0].children[1].children[0].children[0].children[2].children[0]`
    );
    const DateRange3Text = await DateRange3.getText();
    await RPA.Logger.info(`AAAMSの終了日  → ${DateRange3Text}`);
    await RPA.Logger.info(`シートの終了日 → ${WorkData[0][29]}`);
    if (
      DateRange2Text == WorkData[0][28] &&
      DateRange3Text == WorkData[0][29]
    ) {
      await RPA.Logger.info('期間が一致のため次に進みます');
    } else if (
      DateRange2Text != WorkData[0][28] &&
      DateRange3Text == WorkData[0][29]
    ) {
      await RPA.Logger.info('開始日が不一致です');
      await ErrorSetValue(`AJ${Row}`);
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName}!BC${Row}:BC${Row}`,
        values: [['開始日が不一致です']]
      });
      await Start();
    } else if (
      DateRange2Text == WorkData[0][25] &&
      DateRange3Text != WorkData[0][26]
    ) {
      await RPA.Logger.info('終了日が不一致です');
      await ErrorSetValue(`AJ${Row}`);
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName}!BD${Row}:BD${Row}`,
        values: [['終了日が不一致です']]
      });
      await Start();
    } else if (
      DateRange2Text != WorkData[0][25] &&
      DateRange3Text != WorkData[0][26]
    ) {
      await RPA.Logger.info('開始日・終了日が不一致です');
      await ErrorSetValue(`AJ${Row}`);
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName}!BE${Row}:BE${Row}`,
        values: [['開始日・終了日が不一致です']]
      });
      await Start();
    }
    // imp比率(AE列)を入力
    const ImpWeight: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('ads[0].impWeight')[0].children[1].children[0].children[0]`
    );
    await ImpWeight.clear();
    if (WorkData[0][30].length < 1) {
      await RPA.Logger.info('imp比率 の記載がないためスルーします');
    } else {
      await RPA.WebBrowser.sendKeys(ImpWeight, [WorkData[0][30]]);
    }
    // 予約放送枠ID(D列)を入力
    const AdvertisementSlotId: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('ads[0].slotIds[0]')[0].children[0].children[0]`
    );
    await AdvertisementSlotId.clear();
    await RPA.sleep(100);
    if (WorkData[0][3] == 'なし') {
      await RPA.Logger.info('予約放送枠ID が "なし" のためスルーします');
    } else {
      await RPA.WebBrowser.sendKeys(AdvertisementSlotId, [WorkData[0][3]]);
    }
    await RPA.sleep(300);
  }
}

async function AdJudgeError() {
  await RPA.Logger.info('広告作成 エラー判定開始...');
  if (
    WorkData[0][28].length < 1 ||
    WorkData[0][29].length < 1 ||
    (WorkData[0][28].length < 1 && WorkData[0][29].length < 1)
  ) {
    await RPA.Logger.info(`エラー【[広告作成] 配信期間：必須項目です】`);
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AM${Row}:AM${Row}`,
      values: [['【広告作成】必須項目です']]
    });
  }
  if (WorkData[0][30].length < 1) {
    await RPA.Logger.info(`エラー【[広告作成] 比率：${ErrorText4}】`);
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AW${Row}:AW${Row}`,
      values: [[`【広告作成】${ErrorText4}`]]
    });
  }
  const NotAdvertisementSlotId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('ads[0].slotIds[0]')[0].children[1]`
  );
  const NotAdvertisementSlotIdText = await NotAdvertisementSlotId.getText();
  if (NotAdvertisementSlotIdText == '該当する放送枠が存在しません') {
    await RPA.Logger.info(
      `エラー【予約放送枠ID：${NotAdvertisementSlotIdText}】`
    );
    await ErrorSetValue(`AJ${Row}`);
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!AN${Row}:AN${Row}`,
      values: [[`【[広告作成] ${NotAdvertisementSlotIdText}】`]]
    });
  }
  await SheetJudgeError(Row);
}

async function GetAdvertisementId() {
  await RPA.Logger.info('広告 ID発番します');
  const OKButton4: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Modal__footer___aDbwS Modal__full___3Mi9N')[0].children[1]`
  );
  await RPA.WebBrowser.mouseClick(OKButton4);
  await RPA.Logger.info('発番中...');
  await RPA.sleep(5000);
  // 発番した広告IDの最右側の「・・・」をマウスオーバー
  const BalloonMenu = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'Table__menuBtn___1CKo8'
    }),
    15000
  );
  await RPA.WebBrowser.mouseMove(BalloonMenu);
  await RPA.sleep(1000);
  // 「配信を変更する」をクリック
  const ChangeDelivery: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('BalloonMenu__lists___mNMf_ BalloonMenu__arrowTop___222_e')[0].children[3]`
  );
  await RPA.WebBrowser.mouseClick(ChangeDelivery);
  await RPA.sleep(1000);
  const OKButton5: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Modal__footer___aDbwS Modal__mini___cBDF9')[0].children[1]`
  );
  await RPA.WebBrowser.mouseClick(OKButton5);
  await RPA.sleep(8000);
  const AdvertisementId = await RPA.WebBrowser.findElementByClassName(
    'Table__cell___3h99d Table__adId___1Lp24'
  );
  const AdvertisementIdText = await AdvertisementId.getText();
  await RPA.Logger.info(AdvertisementIdText);
  // キャンペーン・広告貼り付けシートのBF列に発番した広告IDを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!BF${Row - 1}:BF${Row - 1}`,
    values: [[AdvertisementIdText]]
  });
  await RPA.Logger.info('広告 作成完了しました');
  // 作業完了を記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!AJ${Row}:AJ${Row}`,
    values: [['完了']]
  });
}

// "エラー"と記載する関数
async function ErrorSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[errortext]]
  });
}

// "必須入力です"と記載する関数
async function RequiredInputSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[ErrorText4]]
  });
}

// "シリーズIDが存在しません"と記載する関数
async function NotSeriesId(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[ErrorText5]]
  });
}

// シートに"エラー"の記載がある場合はスタートに戻る関数
async function SheetJudgeError(Row) {
  // 最新の作業用フラグ(AJ列)を取得
  const SheetJudgeError = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!AJ${Row}:AJ${Row}`
  });
  if (SheetJudgeError[0][0] == errortext) {
    await RPA.Logger.info(
      'シートに "エラー" の記載があるためスタートに戻ります'
    );
    await Start();
  }
}
