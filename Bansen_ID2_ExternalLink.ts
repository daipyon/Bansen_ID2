import RPA from "ts-rpa";
import { WebDriver, By, FileDetector, Key, WebElement } from "selenium-webdriver";
// デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
RPA.Logger.level = 'INFO';

// スプレッドシートIDとシート名
const SSID = process.env.Bansen_ID2_SheetID;
const SSName3 = process.env.Bansen_ID2_SheetName3;
const SSName4 = process.env.Bansen_ID2_SheetName4;
// Abematvのログイン用メールアドレス・パスワードの記載 <<漏洩注意>>
const AbematvID = process.env.Bansen_ID2_AbematvID;
const AbematvPW = process.env.Bansen_ID2_AbematvPW;
// AAAMS(本番環境)のログイン用メールアドレス・パスワード
const AAAMSID = process.env.Bansen_ID2_AAAMSID;
const AAAMSPW = process.env.Bansen_ID2_AAAMSPW;
// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;



const FirstLoginFlag = ['true'];
async function WorkStart(JudgeCampaignId) {
  await RPA.Google.authorize({
    // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: "Bearer",
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  RPA.Logger.info('作業開始...');
  // キャンペーンIDと広告IDを保持する変数
  const IdList = [['','']];
  // 作業対象行とデータを取得
  const WorkData = [];
  const Row = [];
  await GetDataRow(WorkData, Row);

  // キャンペーン名ごとに【外部リンク付き番宣】/【買える】でアカウントを切り替える
  const AccountFlag = ['null'];
  if (WorkData[0][0][0] == '外部リンク付き番宣') {
    AccountFlag[0] = '外部リンク付き番宣';
  }
  if (WorkData[0][0][0] == '買える') {
    AccountFlag[0] = '買える';
  }
  const error = [['エラー']];
  if (WorkData[0][0][0] != '外部リンク付き番宣' && WorkData[0][0][0] != '買える') {
    const ErrorText19 = [['対象アカウントが見つかりませんでした']];
    RPA.Logger.info(`エラー【アカウント：${ErrorText19}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AZ${Row[0]}:AZ${Row[0]}`,values:ErrorText19});
    return Start();
  }

  // AAAMSにログイン
  if (FirstLoginFlag[0] == 'true') {
    await AAAMSLogin(AccountFlag);
  }
  if (FirstLoginFlag[0] == 'false') {
    await AAAMSLogin2(AccountFlag);
  }
  // 一度ログインしたら、以降はログインページをスキップ
  FirstLoginFlag[0] = 'false';

  // キャンペーンIDが発番されていない場合、キャンペーン作成からスタート
  if (JudgeCampaignId.length < 1) {
    RPA.Logger.info('キャンペーン 作成します');
    // キャンペーン作成を押下
    await CampaignStart();

    // 共通項目のデータを入力
    await CommonData(WorkData);

    // 配信先(F列)、配信方式(G列)に入力されている値から処理を分岐
    const Linear = 'リニア';
    const Video = 'ビデオ,タイムシフト';
    const Kpi = 'KPI配信';
    const Shitei = '指定配信';
    // 配信先設定(F列)を選択
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      await CampaignDeliver(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      await CampaignDeliver2(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      await CampaignDeliver2(WorkData);
      // 以下のパターンの場合は作業をスキップし、スタートに戻る
    } else if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Shitei) {
      RPA.Logger.info(`配信先が「${Linear}」、配信方式が「${Shitei}」のためスキップします`);
      const ErrorText = [['エラー', `配信先が「${Linear}」、配信方式が「${Shitei}」のためスキップしました`]];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AH${Row[0]}`,values:ErrorText});
      return Start();
    }

    // 共通項目のエラー判定
    await CommonJudgeError(Row, WorkData);

    // F列が「リニア」、G列が「KPI配信」の場合
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      RPA.Logger.info(`${WorkData[0][0][5]}と${WorkData[0][0][6]}です`);
      // K〜M列、O〜R列を入力
      await Pattern2(Row, WorkData);
    // F列が「ビデオ,タイムシフト」、G列が「KPI配信」の場合
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      // R列、M列、P列、S列を入力
      await Pattern3(Row, WorkData);
    // F列が「ビデオ,タイムシフト」、G列が「指定配信」の場合
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      // M列、P列、S〜T列を入力
      await Pattern4(Row, WorkData);
    }

    // クラスタ(M列)を選択
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      await Cluster(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      await Cluster2(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      await Cluster2(WorkData);
    }

    // 各パターンのエラー判定
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      await Pattern2JudgeError(WorkData, Row);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      await Pattern3JudgeError(WorkData, Row);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      await Pattern4JudgeError(WorkData, Row);
    }

    // キャンペーンIDを発番し、取得
    await GetCampaignId(WorkData, IdList, Row);

    // 広告作成を開始
    await AdvertisementStart(WorkData, Row);

    // 広告作成のエラー判定
    await AdJudgeError(WorkData, Row);

    // 広告作成を発番
    await GetAdvertisementId(IdList, Row);
  }
  // 既にキャンペーンIDが発番されている場合、広告作成からスタート
  if (JudgeCampaignId.length > 1) {
    RPA.Logger.info('広告 作成します');
    await AdvertisementStart2(AccountFlag, WorkData, Row);

    await AdJudgeError2(WorkData, Row);

    await GetAdvertisementId2(Row);
  }
}


async function Start() {
  try {
    await RPA.Google.authorize({
      //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: "Bearer",
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    var count = 0;
    while (0 == 0) {
      // キャンペーン広告貼り付け（外部リンク付き）シートのB〜BB列を取得
      const JudgeData = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName4}!B4:BB3000`});
      // シートのA〜AH列を取得
      const AllData = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!A3:AH3000`});
        RPA.Logger.info('キャンペーンIDです');
        RPA.Logger.info(JudgeData[count][46]);
        if (AllData[count][32] == '初期') { 
          const JudgeCampaignId = JudgeData[count][46];
          await WorkStart(JudgeCampaignId);
        }
      count += 1;
      if (AllData[count][32] == '完了') {
        break;
      }
    }
  } catch (error) {
    RPA.SystemLogger.error(error);
    await RPA.WebBrowser.takeScreenshot();
    // Slackにも通知
    await RPA.Slack.chat.postMessage({
      token: Slack_Token,
      channel: Slack_Channel,
      text: '【番宣 ID発行②（外部リンク）】でエラーが発生しました！'
    });
  }
  RPA.Logger.info('作業を終了します');
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();


async function AAAMSLogin(AccountFlag) {
  await RPA.WebBrowser.get(process.env.AAAMS_Login_URL);
  await RPA.sleep(2000);
  try {
    const AAAMS_loginID_ele = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath:
          '/html/body/div[2]/div/div[2]/form/div/div/div[3]/span/div/div/div/div/div/div/div/div/div[3]/div[1]/div/input'
      }),
      8000
    );
    await RPA.WebBrowser.sendKeys(AAAMS_loginID_ele, [AAAMSID]);
    const AAAMS_loginPW_ele = RPA.WebBrowser.findElementByXPath(
      '/html/body/div[2]/div/div[2]/form/div/div/div[3]/span/div/div/div/div/div/div/div/div/div[3]/div[2]/div/div/input'
    );
    await RPA.WebBrowser.sendKeys(AAAMS_loginPW_ele, [AAAMSPW]);
    const AAAMS_LoginNextButton = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div[2]/div/div[2]/form/div/div/button'
    );
    await RPA.WebBrowser.mouseClick(AAAMS_LoginNextButton);
    await RPA.sleep(3000);
  } catch {
    RPA.Logger.info('ログイン画面をスキップします');
  }

  await RPA.sleep(2000);
  // チャンネル更新画面が出るため待機
  try {
    const ChannelAlart = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div/div/div[6]/div[2]/header/div'
      }),
      5000
    );
    const ChannelAlartButton = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div/div[6]/div[2]/footer/div[2]'
    );
    const AlartText = await ChannelAlart.getText();
    if (AlartText == '下記更新されました。設定を確認してください。') {
      await RPA.WebBrowser.mouseClick(ChannelAlartButton);
      await RPA.sleep(1500);
    }
  } catch {}
  try {
    const Alart = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div/div/div[5]/div[2]/div/p'
      }),
      3000
    );
    const Alartbutton = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div/div[5]/div[2]/footer/div[1]'
    );
    await RPA.WebBrowser.mouseClick(Alartbutton);
    await RPA.sleep(2000);
  } catch {
    RPA.Logger.info('AAAMS アラートが出ませんでしたので次に進みます');
  }
  if (AccountFlag == '外部リンク付き番宣') {
    RPA.Logger.info('外部リンク付き自社広告アカウントを直接呼び出します');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_1);
  }
  if (AccountFlag == '買える') {
    RPA.Logger.info('買えるAbemaTV社アカウントを直接呼び出します');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_2);
  }
  await RPA.sleep(3000);
  // 更新画面をスルー
  try {
    const Koushin = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div[1]/div/div[5]/div[2]/header/div'
      }),
      5000
    );
    const KoushinText = await Koushin.getText();
    if (KoushinText.length > 1) {
      const NextButton01 = await RPA.WebBrowser.findElementByXPath(
        '/html/body/div[1]/div/div[5]/div[2]/footer/div[1]'
      );
      await RPA.WebBrowser.mouseClick(NextButton01);
      await RPA.sleep(1000);
    }
  } catch {
    RPA.Logger.info('更新画面は出ませんでした');
  }
}


async function AAAMSLogin2(AccountFlag) {
  if (AccountFlag == '外部リンク付き番宣') {
    RPA.Logger.info('外部リンク付き自社広告アカウントを直接呼び出します');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_1);
  }
  if (AccountFlag == '買える') {
    RPA.Logger.info('買えるAbemaTV社アカウントを直接呼び出します');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_2);
  }
  await RPA.sleep(2300);
}


async function GetDataRow(WorkData, Row) {
  // 作業用フラグ(AG列)を取得
  const WorkRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG1:AG200`});
  for(let i in WorkRow){
    if (WorkRow[i][0].indexOf('初期') == 0) {
      Row[0] = Number(i) + 1;
      break;
    }
  }
  RPA.Logger.info('この行の作業を開始します → ', Row[0]);
  // シートから作業対象行のデータを取得
  WorkData[0] = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!A${Row[0]}:AD${Row[0]}`});
  RPA.Logger.info(WorkData[0]);
  // AG列に”作業中”と記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:[['作業中']]});
  RPA.Logger.info(`${Row[0]} 行目のステータスを”作業中”に変更しました`);
}


async function CampaignStart() {
  // 左側のキャンペーンをクリック
  const Campaign = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/a[7]');
  await RPA.WebBrowser.mouseClick(Campaign);
  await RPA.sleep(1500);
  // 右上のキャンペーン作成を押下
  const CreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[2]/div[3]/div/header/div/div'}),8000);
  await RPA.WebBrowser.mouseClick(CreateButton);
  await RPA.sleep(3000);
}


// 共通のデータを入力
async function CommonData(WorkData) {
  // キャンペーン名(B列)を入力
  const CampaignName = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[2]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(CampaignName,[WorkData[0][0][1]]);
  // 有効期間をクリック
  const DateRange = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[3]/div[2]/div/div/div[1]/div');
  await RPA.WebBrowser.mouseClick(DateRange);
  // 有効期間：開始(C列)を入力
  const DateRangeStart = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[3]/div[1]/div[1]/input'}), 5000);
  await DateRangeStart.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeStart,[WorkData[0][0][2]]);
  await RPA.sleep(300);
  // 有効期間：終了(D列)を入力
  const DateRangeEnd = await RPA.WebBrowser.findElementByXPath('/html/body/div[3]/div[2]/div[1]/input');
  await DateRangeEnd.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeEnd,[WorkData[0][0][3]]);
  await RPA.sleep(300);
  // 有効期間のOKボタンを押下
  const OKButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[3]/div[3]/div/button[1]');
  await RPA.WebBrowser.mouseClick(OKButton);
  await RPA.sleep(700);
  // キャンペーン種別(E列)を入力
  const CampaignPromo = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[2]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignPromo,[WorkData[0][0][4]]);
  await RPA.WebBrowser.sendKeys(CampaignPromo,[RPA.WebBrowser.Key.ENTER]);
  // 配信方式設定(G列)を入力
  const CampaignPlacement = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[6]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignPlacement,[WorkData[0][0][6]]);
  await RPA.WebBrowser.sendKeys(CampaignPlacement,[RPA.WebBrowser.Key.ENTER]);
}


// 配信先設定(F列)の判定
async function CampaignDeliver(WorkData) {
  RPA.Logger.info('配信先設定 を選択します');
  // 配信先設定の指定がない場合
  if (WorkData[0][0][5].length < 1) {
    const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[4]/div[2]/div/div[2]/div[1]');
    await RPA.WebBrowser.mouseClick(Uncheck);
  } else {
    RPA.Logger.info('配信先設定の指定があるためスルーします');
  }
}


async function CampaignDeliver2(WorkData) {
  RPA.Logger.info('配信先設定 を選択します');
  if (WorkData[0][0][5].length < 1) {
    const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[4]/div[2]/div/div[2]/div[1]');
    await RPA.WebBrowser.mouseClick(Uncheck);
  } else {
    // リニアのチェックを外して、ビデオ・タイムシフトにチェック
    RPA.Logger.info('ビデオ・タイムシフトにチェックします');
    const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[4]/div[2]/div/div[2]/div[1]');
    await RPA.WebBrowser.mouseClick(Uncheck);
    const Check = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[4]/div[2]/div/div[2]/div[2]');
    await RPA.WebBrowser.mouseClick(Check);
    const Check2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[4]/div[2]/div/div[2]/div[3]');
    await RPA.WebBrowser.mouseClick(Check2);
  }
}


// 共通項目のエラー判定
async function CommonJudgeError(Row, WorkData) {
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    return Start();
  }
  RPA.Logger.info('共通項目 エラー判定開始...');
  const error = [['エラー']];
  const DoubleCampaignName = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[2]/div[1]'}),1000);
  const DoubleCampaignNameText = await DoubleCampaignName.getText();
  if (String(DoubleCampaignNameText) == 'キャンペーン名必須\n同じキャンペーン名が既に存在しています') {
    RPA.Logger.info(`エラー【${DoubleCampaignNameText}】`);
    const ErrorText = [['同じキャンペーン名が既に存在しています']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AI${Row[0]}:AI${Row[0]}`,values:ErrorText});
  }
  if (String(DoubleCampaignNameText) == 'キャンペーン名必須\n最大100文字まで入力できます') {
    RPA.Logger.info(`エラー【${DoubleCampaignNameText}】`);
    const ErrorText2 = [['キャンペーン名が65文字以上です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AI${Row[0]}:AI${Row[0]}`,values:ErrorText2});
  }
  if (String(DoubleCampaignNameText) == 'キャンペーン名必須\n必須入力です') {
    RPA.Logger.info(`エラー【${DoubleCampaignNameText}】`);
    const ErrorText3 = [['キャンペーン名：必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AI${Row[0]}:AI${Row[0]}`,values:ErrorText3});
  }
  if (WorkData[0][0][2].length < 1 || WorkData[0][0][3].length < 1 || WorkData[0][0][2].length < 1 && WorkData[0][0][3].length < 1) {
    const ErrorText4 = [['必須項目です']];
    RPA.Logger.info(`エラー【有効期間：${ErrorText4}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AJ${Row[0]}:AJ${Row[0]}`,values:ErrorText4});
  }
  if (WorkData[0][0][5].length < 1) {
    const ErrorText6 = [['必須入力です']];
    RPA.Logger.info(`エラー【配信先設定：${ErrorText6}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AK${Row[0]}:AK${Row[0]}`,values:ErrorText6});
  }
  if (WorkData[0][0][6].length < 1) {
    const ErrorText7 = [['必須入力です']];
    RPA.Logger.info(`エラー【配信方式設定：${ErrorText7}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AL${Row[0]}:AL${Row[0]}`,values:ErrorText7});
  }
}


// H〜L列、M列、O〜P列を入力
async function Pattern2(Row, WorkData) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    return Start();
  }
  RPA.Logger.info('パターン 2 です');
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[2]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[WorkData[0][0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[RPA.WebBrowser.Key.ENTER]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[3]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[WorkData[0][0][8]]);
  await CampaignFilterTempId.sendKeys(Key.SPACE);
  await RPA.sleep(500);
  await CampaignFilterTempId.sendKeys(Key.BACK_SPACE);
  await RPA.sleep(1000);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName('Select-option');
  const FilterTempValueText = await Promise.all(FirstFilterTempValue.map(async elm => (await elm.getText())));
  RPA.Logger.info('配信先フィルタテンプレート一覧 → '+ FilterTempValueText);
  for (let i in FilterTempValueText) {
    if (WorkData[0][0][8] == FilterTempValueText[i]) {
      RPA.Logger.info(`${FilterTempValueText[i]} 一致しました`);
      // const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(`/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/div/div[${Number(i) + Number(1)}]`);
      await RPA.sleep(1000);
      // await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[RPA.WebBrowser.Key.ENTER]);
      FilterTempFlag[0] = false;
      break;
    }
    await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[RPA.WebBrowser.Key.DOWN]);
  }
  await RPA.sleep(1000);
  // 取得した値を分割
  const ValuesJ = WorkData[0][0][9].split(',');
  await RPA.Logger.info(ValuesJ);
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[4]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  for (let i in ValuesJ) {
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[ValuesJ[i]]);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[RPA.WebBrowser.Key.ENTER]);
  }
  // 隣接許容レベル・広告(K列)を入力
  if (WorkData[0][0][10].length < 1) {
    const CroseButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[2]');
    await RPA.WebBrowser.mouseClick(CroseButton);
  } else {
    const CampaignAdjacentAcceptableAd = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd,[WorkData[0][0][10]]);
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd,[RPA.WebBrowser.Key.ENTER]);
  }
  // 隣接許容レベル・クリエイティブ(L列)を選択
  if (WorkData[0][0][11].length < 1) {
    const CroseButton2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[2]');
    await RPA.WebBrowser.mouseClick(CroseButton2);
  } else {
    const CampaignAdjacentAcceptableCreative = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative,[WorkData[0][0][11]]);
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative,[RPA.WebBrowser.Key.ENTER]);
  }
  // 指定シリーズNG(M列)を入力
  const CampaignSeriesNg = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[1]/div[2]/textarea');
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[WorkData[0][0][12]]);
  await CampaignSeriesNg.sendKeys(Key.SPACE);
  await RPA.sleep(1000);
  await CampaignSeriesNg.sendKeys(Key.BACK_SPACE);
  // 優先度(O列)を入力
  const GoalPriority: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('goal.priority')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
  );
  await RPA.WebBrowser.sendKeys(GoalPriority,[WorkData[0][0][14]]);
  await RPA.WebBrowser.sendKeys(GoalPriority,[RPA.WebBrowser.Key.ENTER]);
  // 総目標imp(P列)を入力
  const GoalGimp: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('goal.gimp')[0].children[1].children[0].children[0]`
  );
  await RPA.WebBrowser.sendKeys(GoalGimp,[WorkData[0][0][15]]);
  await RPA.WebBrowser.sendKeys(GoalGimp,[RPA.WebBrowser.Key.ENTER]);
  //「FQキャップ」に「3」と入力（2020/07/09 追加）
  const FqCap: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FieldText__inputWrap___24SZl FieldText__unitWrap___2SOu6')[2].children[0]`
  );
  await FqCap.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(FqCap, [`3`]);
  await RPA.sleep(300);
}


// H〜J列、M列、P列、R〜S列を入力
async function Pattern3(Row, WorkData) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    return Start();
  }
  RPA.Logger.info('パターン 3 です');
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[2]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[WorkData[0][0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[RPA.WebBrowser.Key.ENTER]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[3]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[WorkData[0][0][8]]);
  await CampaignFilterTempId.sendKeys(Key.SPACE);
  await RPA.sleep(500);
  await CampaignFilterTempId.sendKeys(Key.BACK_SPACE);
  await RPA.sleep(1000);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName('Select-option');
  const FilterTempValueText = await Promise.all(FirstFilterTempValue.map(async elm => (await elm.getText())));
  RPA.Logger.info('配信先フィルタテンプレート一覧 → '+ FilterTempValueText);
  for (let i in FilterTempValueText) {
    if (WorkData[0][0][8] == FilterTempValueText[i]) {
      RPA.Logger.info(`${FilterTempValueText[i]} 一致しました`);
      // const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(`/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/div/div[${Number(i) + Number(1)}]`);
      await RPA.sleep(1000);
      // await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[RPA.WebBrowser.Key.ENTER]);
      FilterTempFlag[0] = false;
      break;
    }
    await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[RPA.WebBrowser.Key.DOWN]);
  }
  await RPA.sleep(1000);
  // 取得した値を分割
  const ValuesJ = WorkData[0][0][9].split(',');
  await RPA.Logger.info(ValuesJ);
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[4]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  for (let i in ValuesJ) {
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[ValuesJ[i]]);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[RPA.WebBrowser.Key.ENTER]);
  }
  // 指定シリーズNG(M列)を入力
  const CampaignSeriesNg = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[1]/div[2]/textarea');
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[WorkData[0][0][12]]);
  await CampaignSeriesNg.sendKeys(Key.SPACE);
  await RPA.sleep(1000);
  await CampaignSeriesNg.sendKeys(Key.BACK_SPACE);
  // 総目標imp(P列)を入力
  const GoalGimp: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('goal.gimp')[0].children[1].children[0].children[0]`
  );
  await RPA.WebBrowser.sendKeys(GoalGimp,[WorkData[0][0][15]]);
  await RPA.WebBrowser.sendKeys(GoalGimp,[RPA.WebBrowser.Key.ENTER]);
  await RPA.sleep(1000);
  // 訴求シリーズID(R列)を入力
  const CampaignPromoSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.promoSeriesId[0]')[0].children[0].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CampaignPromoSeriesId,[WorkData[0][0][17]]);
  // カスタムオーディエンス(S列)を入力
  const CustomAudience: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.customAudienceString.id')[0].children[1].children[0].children[0].children[0].children[0].children[0].children[0].children[1].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CustomAudience,[WorkData[0][0][18]]);
  await RPA.WebBrowser.sendKeys(CustomAudience,[RPA.WebBrowser.Key.ENTER]);
  //「FQキャップ」に「3」と入力（2020/07/09 追加）
  const FqCap: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FieldText__inputWrap___24SZl FieldText__unitWrap___2SOu6')[2].children[0]`
  );
  await FqCap.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(FqCap, [`3`]);
  await RPA.sleep(300);
}


// H〜J列、P列、R列、Q列を入力
async function Pattern4(Row, WorkData) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    return Start();
  }
  RPA.Logger.info('パターン 4 です');
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[2]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[WorkData[0][0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[RPA.WebBrowser.Key.ENTER]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[3]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[WorkData[0][0][8]]);
  await CampaignFilterTempId.sendKeys(Key.SPACE);
  await RPA.sleep(500);
  await CampaignFilterTempId.sendKeys(Key.BACK_SPACE);
  await RPA.sleep(1000);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName('Select-option');
  const FilterTempValueText = await Promise.all(FirstFilterTempValue.map(async elm => (await elm.getText())));
  RPA.Logger.info('配信先フィルタテンプレート一覧 → '+ FilterTempValueText);
  for (let i in FilterTempValueText) {
    if (WorkData[0][0][8] == FilterTempValueText[i]) {
      RPA.Logger.info(`${FilterTempValueText[i]} 一致しました`);
      // const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(`/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/div/div[${Number(i) + Number(1)}]`);
      await RPA.sleep(1000);
      // await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[RPA.WebBrowser.Key.ENTER]);
      FilterTempFlag[0] = false;
      break;
    }
    await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[RPA.WebBrowser.Key.DOWN]);
  }
  await RPA.sleep(1000);
  // 取得した値を分割
  const ValuesJ = WorkData[0][0][9].split(',');
  await RPA.Logger.info(ValuesJ);
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[4]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  for (let i in ValuesJ) {
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[ValuesJ[i]]);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[RPA.WebBrowser.Key.ENTER]);
  }
  // 総目標imp(P列)を入力
  const GoalGimp: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('goal.gimp')[0].children[1].children[0].children[0]`
  );
  await RPA.WebBrowser.sendKeys(GoalGimp,[WorkData[0][0][15]]);
  await RPA.WebBrowser.sendKeys(GoalGimp,[RPA.WebBrowser.Key.ENTER]);
  await RPA.sleep(1000);
  // 訴求シリーズID(R列)を入力
  const CampaignPromoSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.promoSeriesId[0]')[0].children[0].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CampaignPromoSeriesId,[WorkData[0][0][17]]);
  // 指定シリーズID(Q列)を入力
  const CampaignSpecifiedSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FieldText__textarea___3vdMx')[1]`
  );
  const Text1 = await encodeURI(WorkData[0][0][16]);
  await RPA.WebBrowser.driver.executeScript(
    // `document.querySelectorAll("#reactroot > div > div.Modal__overlay___3CdlJ.Modal__full___3Mi9N.Modal__opened___-zyiB > div.Modal__modal___15eHS.Modal__full___3Mi9N > div.Modal__body___2lTdW.Modal__full___3Mi9N > div > form > div > div:nth-child(4) > div.FieldGroup__container___3fkr0 > div > div:nth-child(10) > div.Field__value___3-hi8 > div > div.Field__container___2rw9q.Field__column___2-TfA.Field__labelInline___2GdCv.Field__noMargin___1oAyZ.Field__alignNone___1dlxQ > div.Field__value___3-hi8 > textarea")[0].value = decodeURI("${Text1}");`
    `document.getElementsByClassName('FieldText__textarea___3vdMx')[1].value = decodeURI("${Text1}");`
  );
  await RPA.WebBrowser.sendKeys(CampaignSpecifiedSeriesId, [RPA.WebBrowser.Key.ENTER]);
  //「FQキャップ」に「3」と入力（2020/07/09 追加）
  const FqCap: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FieldText__inputWrap___24SZl FieldText__unitWrap___2SOu6')[2].children[0]`
  );
  await FqCap.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(FqCap, [`3`]);
  await RPA.sleep(300);
}


// クラスタ(N列)を選択
async function Cluster(WorkData) {
  RPA.Logger.info('クラスタ を選択します');
  // セルの値を分割
  const ValuesN = WorkData[0][0][13].split(',');
  // クラスタ・全て選択をクリック
  const All = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[2]/div/div[1]');
  await RPA.WebBrowser.mouseClick(All);
  for (var i = 0; i <= ValuesN.length - 1; i++) {
    // クラスタ・F1を取得
    const F1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[2]/div/div[2]/div[1]');
    const CheckF1 = await F1.getText();
    if (CheckF1 == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(F1);
    }
    // クラスタ・F2以上を取得
    const F2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[2]/div/div[2]/div[2]');
    const F2Ijyou = await F2.getText();
    // "以上"という文字を削除
    const CheckF2 = F2Ijyou.slice(0, -2);
    if (CheckF2　== ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(F2);
    }
    // クラスタ・M1を取得
    const M1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[2]/div/div[2]/div[3]');
    const CheckM1 = await M1.getText();
    if (CheckM1 == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(M1);
    }
    // クラスタ・M2以上を取得
    const M2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[2]/div/div[2]/div[4]');
    const M2Ijyou = await M2.getText();
    // "以上"という文字を削除
    const CheckM2 = M2Ijyou.slice(0, -2);
    if (CheckM2 == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(M2);
    }
    // クラスタ・teenを取得
    const Teen = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[2]/div/div[2]/div[5]');
    const CheckTeen = await Teen.getText();
    if (CheckTeen == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(Teen);
    }
    // クラスタ・otherを取得
    const Other = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[2]/div/div[2]/div[6]');
    const CheckOther = await Other.getText();
    if (CheckOther == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(Other);
    }
    const CheckAll = await All.getText();
    if (CheckAll == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(All);
    }
  }
}


async function Cluster2(WorkData) {
  RPA.Logger.info('クラスタ を選択します');
  // セルの値を分割
  const ValuesN =WorkData[0][0][13].split(',');
  // クラスタ・全て選択をクリック
  const All = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[1]');
  await RPA.WebBrowser.mouseClick(All);
  for (var i = 0; i <= ValuesN.length - 1; i++) {
    // クラスタ・F1を取得
    const F1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[2]/div[1]');
    const CheckF1 = await F1.getText();
    if (CheckF1 == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(F1);
    }
    // クラスタ・F2以上を取得
    const F2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[2]/div[2]');
    const F2Ijyou = await F2.getText();
    // "以上"という文字を削除
    const CheckF2 = F2Ijyou.slice(0, -2);
    if (CheckF2　== ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(F2);
    }
    // クラスタ・M1を取得
    const M1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[2]/div[3]');
    const CheckM1 = await M1.getText();
    if (CheckM1 == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(M1);
    }
    // クラスタ・M2以上を取得
    const M2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[2]/div[4]');
    const M2Ijyou = await M2.getText();
    // "以上"という文字を削除
    const CheckM2 = M2Ijyou.slice(0, -2);
    if (CheckM2 == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(M2);
    }
    // クラスタ・teenを取得
    const Teen = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[2]/div[5]');
    const CheckTeen = await Teen.getText();
    if (CheckTeen == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(Teen);
    }
    // クラスタ・otherを取得
    const Other = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[2]/div[6]');
    const CheckOther = await Other.getText();
    if (CheckOther == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(Other);
    }
    const CheckAll = await All.getText();
    if (CheckAll == ValuesN[i]) {
      await RPA.WebBrowser.mouseClick(All);
    }
  }
}


async function Pattern2JudgeError(WorkData, Row) {
  RPA.Logger.info('パターン 2 エラー判定開始...');
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  // const FieldError4 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[1]'}),1000);
  // const FieldErrorText4 = await FieldError4.getText();
  // if (String(FieldErrorText4) == '配信先曜日×時間帯テンプレート必須\n必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText4}】`);
  //   const ErrorText8 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AM${Row[0]}:AM${Row[0]}`,values:ErrorText8});
  // }
  // const FieldError5 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[1]'}),1000);
  // const FieldErrorText5 = await FieldError5.getText();
  // if (String(FieldErrorText5) == 'リニア配信先フィルタテンプレート必須\n必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText5}】`);
  //   const ErrorText9 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AN${Row[0]}:AN${Row[0]}`,values:ErrorText9});
  // }
  // const FieldError6 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[1]/div[1]'}),1000);
  // const FieldErrorText6 = await FieldError6.getText();
  // if (String(FieldErrorText6) == '広告必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText6}】`);
  //   const ErrorText10 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AO${Row[0]}:AO${Row[0]}`,values:ErrorText10});
  // }
  // const FieldError7 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[2]/div[1]'}),1000);
  // const FieldErrorText7 = await FieldError7.getText();
  // if (String(FieldErrorText7) == 'クリエイティブ必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText7}】`);
  //   const ErrorText11 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AP${Row[0]}:AP${Row[0]}`,values:ErrorText11});
  // }
  if (WorkData[0][0][12].length < 1) {
    RPA.Logger.info('指定シリーズNGの記載がないためスルーします');
  } else {
    const NotCampaignSeriesNg = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[2]/div[2]/div/div[2]'}),1000);
    const NotCampaignSeriesNgText = await NotCampaignSeriesNg.getText();
    if (String(NotCampaignSeriesNgText) == 'シリーズIDが存在しません') {
      RPA.Logger.info(`エラー【${NotCampaignSeriesNgText}】`);
      const ErrorText12 = [['シリーズIDが存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AQ${Row[0]}:AQ${Row[0]}`,values:ErrorText12});
    }
  }
  // if (WorkData[0][0][13].length < 1) {
  //   const ErrorText13 = [['必須入力です']];
  //   RPA.Logger.info(`エラー【クラスタ：${ErrorText13}】`);
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AR${Row[0]}:AR${Row[0]}`,values:ErrorText13});
  // }
  if (WorkData[0][0][14].length < 1) {
    const ErrorText15 = [['必須入力です']];
    RPA.Logger.info(`エラー【優先度：${ErrorText15}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AT${Row[0]}:AT${Row[0]}`,values:ErrorText15});
  }
  const FieldError11 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[9]/div[1]'}),1000);
  const FieldErrorText11 = await FieldError11.getText();
  if (String(FieldErrorText11) == '総目標imp必須必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText11}】`);
    const ErrorText16 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AU${Row[0]}:AU${Row[0]}`,values:ErrorText16});
  }
}


async function Pattern3JudgeError(WorkData, Row) {
  RPA.Logger.info('パターン 3 エラー判定開始...');
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  // const FieldError4 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[1]'}),1000);
  // const FieldErrorText4 = await FieldError4.getText();
  // if (String(FieldErrorText4) == '配信先曜日×時間帯テンプレート必須\n必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText4}】`);
  //   const ErrorText8 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AM${Row[0]}:AM${Row[0]}`,values:ErrorText8});
  // }
  // const FieldError5 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[1]'}),1000);
  // const FieldErrorText5 = await FieldError5.getText();
  // if (String(FieldErrorText5) == 'リニア配信先フィルタテンプレート必須\n必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText5}】`);
  //   const ErrorText9 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AN${Row[0]}:AN${Row[0]}`,values:ErrorText9});
  // }
  if (WorkData[0][0][12].length < 1) {
    RPA.Logger.info('指定シリーズNG の記載がないためスルーします');
  } else {
    const NotCampaignSeriesNg = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[2]/div[2]/div/div[2]'}),1000);
    const NotCampaignSeriesNgText = await NotCampaignSeriesNg.getText();
    if (String(NotCampaignSeriesNgText) == 'シリーズIDが存在しません') {
      RPA.Logger.info(`エラー【${NotCampaignSeriesNgText}】`);
      const ErrorText12 = [['シリーズIDが存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AQ${Row[0]}:AQ${Row[0]}`,values:ErrorText12});
    }
  }
  // if (WorkData[0][0][13].length < 1) {
  //   const ErrorText13 = [['必須入力です']];
  //   RPA.Logger.info(`エラー【クラスタ：${ErrorText13}】`);
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AR${Row[0]}:AR${Row[0]}`,values:ErrorText13});
  // }
  const FieldError11 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[1]'}),1000);
  const FieldErrorText11 = await FieldError11.getText();
  if (String(FieldErrorText11) == '総目標imp必須必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText11}】`);
    const ErrorText16 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AU${Row[0]}:AU${Row[0]}`,values:ErrorText16});
  }
  const NotCampaignPromoSeriesId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[8]/div[2]/div/div[2]'}),1000);
  const NotCampaignPromoSeriesIdText = await NotCampaignPromoSeriesId.getText();
  if (String(NotCampaignPromoSeriesIdText) == 'シリーズIDが存在しません') {
    RPA.Logger.info(`エラー【${NotCampaignPromoSeriesIdText}】`);
    const ErrorText18 = [['シリーズIDが存在しません']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AW${Row[0]}:AW${Row[0]}`,values:ErrorText18});
  }
}


async function Pattern4JudgeError(WorkData, Row) {
  RPA.Logger.info('パターン 4 エラー判定開始...');
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  // const FieldError4 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[1]'}),1000);
  // const FieldErrorText4 = await FieldError4.getText();
  // if (String(FieldErrorText4) == '配信先曜日×時間帯テンプレート必須\n必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText4}】`);
  //   const ErrorText8 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AM${Row[0]}:AM${Row[0]}`,values:ErrorText8});
  // }
  // const FieldError5 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[1]'}),1000);
  // const FieldErrorText5 = await FieldError5.getText();
  // if (String(FieldErrorText5) == 'リニア配信先フィルタテンプレート必須\n必須入力です') {
  //   RPA.Logger.info(`エラー【${FieldErrorText5}】`);
  //   const ErrorText9 = [['必須入力です']];
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AN${Row[0]}:AN${Row[0]}`,values:ErrorText9});
  // }
  // if (WorkData[0][0][13].length < 1) {
  //   const ErrorText13 = [['必須入力です']];
  //   RPA.Logger.info(`エラー【クラスタ：${ErrorText13}】`);
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
  //   await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AR${Row[0]}:AR${Row[0]}`,values:ErrorText13});
  // }
  const FieldError11 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[7]/div[1]'}),1000);
  const FieldErrorText11 = await FieldError11.getText();
  if (String(FieldErrorText11) == '総目標imp必須必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText11}】`);
    const ErrorText16 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AU${Row[0]}:AU${Row[0]}`,values:ErrorText16});
  }
  if (WorkData[0][0][16].length < 1) {
    RPA.Logger.info('エラー：指定シリーズID の記載がありません');
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AV${Row[0]}:AV${Row[0]}`,values:[['IDの記載がありません']]});
  } else {
    const NotCampaignSpecifiedSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName('campaign.vodSeriesIds[0]')[0].children[1]`
    );
    const NotCampaignSpecifiedSeriesIdText = await NotCampaignSpecifiedSeriesId.getText();
    if (String(NotCampaignSpecifiedSeriesIdText) == 'シリーズIDが存在しません') {
      RPA.Logger.info(`エラー【${NotCampaignSpecifiedSeriesId}】`);
      const ErrorText17 = [['シリーズIDが存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AV${Row[0]}:AV${Row[0]}`,values:ErrorText17});
    }
  }
  const NotCampaignPromoSeriesId: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName('campaign.promoSeriesId[0]')[0].children[1]`
  );
  const NotCampaignPromoSeriesIdText = await NotCampaignPromoSeriesId.getText();
  if (String(NotCampaignPromoSeriesIdText) == 'シリーズIDが存在しません') {
    RPA.Logger.info(`エラー【${NotCampaignPromoSeriesIdText}】`);
    const ErrorText18 = [['シリーズIDが存在しません']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AW${Row[0]}:AW${Row[0]}`,values:ErrorText18});
  }
}


// キャンペーンIDを発番
async function GetCampaignId(WorkData, IdList, Row) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    return Start();
  }
  // OKボタンをクリック
  const OKButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath: '//*[@id="reactroot"]/div/div[5]/div[2]/footer/div[2]'}), 5000);
  await RPA.WebBrowser.mouseClick(OKButton);
  RPA.Logger.info('キャンペーン ID発番中...');
  await RPA.sleep(5000);
  // キャンペーン名が一致するか判定
  for (var i = 1; i <= 15; i++) {
    const CampaignName = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath: `/html/body/div[1]/div/div[2]/div[3]/div/table/tbody/tr[${Number(i)}]/td[3]`}), 5000);
    const CampaignNameText = await CampaignName.getText();
    RPA.Logger.info('取得したキャンペーン名　　　　　　　　　→　' + CampaignNameText);
    if (CampaignNameText == WorkData[0][0][1]) {
      await RPA.WebBrowser.mouseClick(CampaignName);
      RPA.Logger.info('現在保持しているデータのキャンペーン名　→　' + WorkData[0][0][1]);
      RPA.Logger.info('キャンペーン名が一致しましたので、IDを取得します');
      break;
    }
  }
  await RPA.sleep(5000);
  // 発番したキャンペーンIDを取得し、キャンペーン広告貼り付け（外部リンク付き）シートに記載
  const CampaignId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath: '/html/body/div/div/div[2]/div[3]/div/div[1]/div[2]/div[1]/div[2]'}), 5000);
  IdList[0][0] = await CampaignId.getText();
  RPA.Logger.info(IdList);
  // キャンペーン名が同じものがあればその行にもキャンペーンIDを記載
  const JudgeCampaignName = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName4}!AB4:AB3000`});
  for (let i in JudgeCampaignName) {
    if (JudgeCampaignName[i][0] == WorkData[0][0][1]) {
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName4}!AV${Number(i) + 4}:AW${Number(i) + 4}`,values:IdList});
    }
  }
  RPA.Logger.info('キャンペーン 作成完了しました');
}


async function AdvertisementStart(WorkData, Row) {
  // 右下の広告作成を押下
  const CreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath: '/html/body/div/div/div[2]/div[3]/div/div[2]/div[2]/div'}), 5000);
  await RPA.WebBrowser.mouseClick(CreateButton);
  await RPA.sleep(3000);
  // クリエイティブID(U列)を入力
  const CreativeId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath: '/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[1]/div[2]/div/input'}), 5000);
  // クリエイティブIDがない場合、作業をスキップしてスタートに戻る
  const error = [['エラー']];
  if (WorkData[0][0][20].length < 1) {
    const ErrorText = [['シートにIDの記載がありません']];
    RPA.Logger.info(`エラー【${ErrorText}】` + ' 作業をスキップします');
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:ErrorText});
    return Start();
  } else {
    await RPA.WebBrowser.sendKeys(CreativeId,[WorkData[0][0][20]]);
    await RPA.sleep(5000);
    // エラーが起きた場合、作業をスキップしてスタートに戻る
    try {
      var ApplyFlag = ['false'];
      await RPA.sleep(300);
      const NotCreativeId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath: '/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/div'}), 1000);
      const NotCreativeIdText = await NotCreativeId.getText();
      if (String(NotCreativeIdText) == '該当する項目が存在しません') {
        RPA.Logger.info(`エラー【${NotCreativeIdText}】` + ' 作業をスキップします');
        const ErrorText2 = [['該当する項目が存在しません']];
        await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
        await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:ErrorText2});
        return Start();
      }
    } catch {
      ApplyFlag[0] = 'true';
      RPA.Logger.info('次の処理に進みます');
    }
    // クリエイティブ選択の「選択」をクリック
    const SelectCreative = await RPA.WebBrowser.findElementByXPath('//*[@id="reactroot"]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[1]/div');
    await RPA.WebBrowser.mouseClick(SelectCreative);
    // 配信期間をクリック
    const DateRange = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[2]/div[2]/div/div/div[1]/div');
    await RPA.WebBrowser.mouseClick(DateRange);
    // 配信期間：開始(C列)を入力
    const DateRangeStart = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[1]/div[1]/input');
    await DateRangeStart.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeStart,[WorkData[0][0][2]]);
    await RPA.sleep(300);
    // 配信期間：終了(D列)を入力
    const DateRangeEnd = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[2]/div[1]/input');
    await DateRangeEnd.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeEnd,[WorkData[0][0][3]]);
    await RPA.sleep(300);
    // 配信期間のOKボタンを押下
    const OKButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[3]/div/button[1]');
    await RPA.WebBrowser.mouseClick(OKButton);
    await RPA.sleep(1000);
    // 一度入力されている開始・終了期間をそれぞれ取得
    const DateRangeText = await DateRange.getText();
    RPA.Logger.info('AAAMSの開始日  → ' + DateRangeText);
    RPA.Logger.info('シートの開始日 → ' + WorkData[0][0][2] + ':00');
    const DateRange2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div');
    const DateRangeText2 = await DateRange2.getText();
    RPA.Logger.info('AAAMSの終了日  → ' + DateRangeText2);
    RPA.Logger.info('シートの終了日 → ' + WorkData[0][0][3] + ':00');
    if (DateRangeText == WorkData[0][0][2] + ':00' && DateRangeText2 == WorkData[0][0][3] + ':00') {
      RPA.Logger.info('期間が一致のため次に進みます');
    } else if (DateRangeText != WorkData[0][0][2] + ':00' && DateRangeText2 == WorkData[0][0][3] + ':00'){
      const NotMotchDateRangeStart = [['開始日が不一致です']];
      RPA.Logger.info(NotMotchDateRangeStart[0][0]);
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:NotMotchDateRangeStart});
      return Start();
    } else if (DateRangeText == WorkData[0][0][2] + ':00' && DateRangeText2 != WorkData[0][0][3] + ':00'){
      const NotMotchDateRangeEnd = [['終了日が不一致です']];
      RPA.Logger.info(NotMotchDateRangeEnd[0][0]);
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:NotMotchDateRangeEnd});
      return Start();
    } else if (DateRangeText != WorkData[0][0][2] + ':00' && DateRangeText2 != WorkData[0][0][3] + ':00'){
      const NotMotchDateRange = [['開始日・終了日が不一致です']];
      RPA.Logger.info(NotMotchDateRange[0][0]);
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:NotMotchDateRange});
      return Start();
    }
    // imp比率(AA列)を入力
    if (WorkData[0][0][26].length < 1) {
      const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
      await ImpWeight.clear();
    } else {
      const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
      await ImpWeight.clear();
      await RPA.WebBrowser.sendKeys(ImpWeight,[WorkData[0][0][26]]);
    }
    // テキスト(AB列)を入力
    const Text = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
    if (WorkData[0][0][27] == '広告なしで ビデオをみる') {
      const WithoutAd = WorkData[0][0][27].slice(0, -6);
      await RPA.WebBrowser.sendKeys(Text,[WithoutAd]);
      Text.sendKeys(Key.ENTER);
    } else {
      await RPA.WebBrowser.sendKeys(Text,[WorkData[0][0][27]]);
      Text.sendKeys(Key.ENTER);
    }
    // URL(AC列)を入力
    const Url = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[2]/div[2]/div/input');
    if (WorkData[0][0][28].length < 1) {
      await RPA.WebBrowser.mouseClick(Url);
    } else {
      await RPA.WebBrowser.sendKeys(Url,[WorkData[0][0][28]]);
      await RPA.sleep(3000);
      await Url.sendKeys(Key.SPACE);
      await RPA.sleep(3000);
      await Url.sendKeys(Key.BACK_SPACE);
    }
    await RPA.sleep(3000);
    // URL(AC列)のエラーを判定するため、申込目標impをクリック
    const ImpRequest = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[7]/div[2]/div/input');
    await RPA.WebBrowser.mouseClick(ImpRequest);
    await RPA.sleep(3000);
    // エラーが出た場合は再入力
    for (var i = 0; i < 10; i++) {
      try {
        const FraudulentUrl = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[2]/div[1]/span'}),1000);
        const FraudulentUrlText = await FraudulentUrl.getText();
        RPA.Logger.info(FraudulentUrlText);
        if (String(FraudulentUrlText) == 'リンク先が不正です') {
          RPA.Logger.info('エラー文が出ましたので再入力します');
          await Url.clear();
          await RPA.sleep(100);
          await RPA.WebBrowser.sendKeys(Url,[WorkData[0][0][28]]);
          await RPA.sleep(3000);
          await Url.sendKeys(Key.SPACE);
          await RPA.sleep(3000);
          await Url.sendKeys(Key.BACK_SPACE);
          await RPA.sleep(3000);
          await RPA.WebBrowser.mouseClick(ImpRequest);
          await RPA.sleep(3000);
        } else {
          break;
        }
      } catch {
        RPA.Logger.info('エラー文が出ませんでしたので次に進みます');
        break;
      }
    }
  }
}


async function AdJudgeError(WorkData, Row) {
  RPA.Logger.info('広告作成 エラー判定開始...');
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  if (WorkData[0][0][2].length < 1 || WorkData[0][0][3].length < 1 || WorkData[0][0][2].length < 1 && WorkData[0][0][3].length < 1) {
    const ErrorText3 = [['【広告作成】必須項目です']];
    RPA.Logger.info(`エラー【${ErrorText3}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AJ${Row[0]}:AJ${Row[0]}`,values:ErrorText3});
  }
  if (WorkData[0][0][26].length < 1) {
    const ErrorText4 = [['【広告作成】必須入力です']];
    RPA.Logger.info(`エラー【${ErrorText4}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AS${Row[0]}:AS${Row[0]}`,values:ErrorText4});
  }
  if (WorkData[0][0][27].length < 1) {
    const ErrorText5 = [['テキストを選択してください']];
    RPA.Logger.info(`エラー【${ErrorText5}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AX${Row[0]}:AX${Row[0]}`,values:ErrorText5});
  }
  if (WorkData[0][0][28].length < 1) {
    const ErrorText6 = [['URLを入力してください']];
    RPA.Logger.info(`エラー【${ErrorText6}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
  const NotUrl = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[2]/div[1]'}),1000);
  const NotUrlText = await NotUrl.getText();
  if (String(NotUrlText) == 'URLURLを入力してください') {
    RPA.Logger.info(`エラー【${NotUrlText}】`);
    const ErrorText6 = [['URLを入力してください']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
  if (String(NotUrlText) == 'URLURLを正しく入力してください') {
    RPA.Logger.info(`エラー【${NotUrlText}】`);
    const ErrorText6 = [['URLを正しく入力してください']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
  if (String(NotUrlText) == 'URLリンク先が不正です') {
    RPA.Logger.info(`エラー【${NotUrlText}】`);
    const ErrorText6 = [['リンク先が不正です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
}


async function GetAdvertisementId(IdList, Row) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    return Start();
  }
  // OKボタンをクリック
  const OKButton2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton2);
  RPA.Logger.info('広告 ID発番中...');
  await RPA.sleep(5000);
  // 発番した広告IDの最右側の「・・・」をマウスオーバー
  const BalloonMenu = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[10]/div'
      }),
      5000
    );
  await RPA.WebBrowser.mouseMove(BalloonMenu);
  await RPA.sleep(1000);
  // 「配信を変更する」をクリック
  const ChangeDelivery = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[3]/div/div[4]');
  await RPA.WebBrowser.mouseClick(ChangeDelivery);
  await RPA.sleep(1000);
  const OKButton3 = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton3);
  await RPA.sleep(8000);
  const AdvertisementId = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[1]');
  IdList[0][1] = await AdvertisementId.getText();
  RPA.Logger.info(IdList);
  // キャンペーン広告貼り付け（外部リンク付き）シートのAV〜AW列に発番した広告IDを記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName4}!AV${Row[0] + 1}:AW${Row[0] + 1}`,values:IdList});
  RPA.Logger.info('広告 作成完了しました');
  // 作業完了を記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:[['完了']]});
}


async function AdvertisementStart2(AccountFlag, WorkData, Row) {
  const LoadingFlag = [];
  LoadingFlag[0] = false;
  while (LoadingFlag[0] == false) {
    const JudgeCampaignId = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName4}!AV${Row[0] + 1}:AV${Row[0] + 1}`});
    // キャンペーンIDで検索
    const SearchByCampaignId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/div[1]/form/div/div[5]/div[2]/div/input');
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(SearchByCampaignId, JudgeCampaignId[0]);
    await RPA.sleep(5000);
    try {
      const Loading = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/div[2]/div/p[1]');
      const LoadingText = await Loading.getText();
      if (LoadingText == '読み込み中...') {
        LoadingFlag[0] = false;
        RPA.Logger.info('検索結果が出ないため、ブラウザを更新し再検索します');
        await RPA.WebBrowser.refresh();
      }
      if (AccountFlag == '外部リンク付き番宣') {
        RPA.Logger.info('外部リンク付き自社広告アカウントを直接呼び出します');
        await RPA.WebBrowser.get(process.env.AAAMS_Account_1);
      }
      if (AccountFlag == '買える') {
        RPA.Logger.info('買えるAbemaTV社アカウントを直接呼び出しします');
        await RPA.WebBrowser.get(process.env.AAAMS_Account_2);
      }
      await RPA.sleep(3000);
      // 代理店の箇所に「外部リンク付き番宣」もしくは「【コマース】買えるAbemaTV」の文字が入っていない場合は再更新
      const AgencyName = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/div[1]/form/div/div[1]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[1]/div[1]');
      const AgencyNameText = await AgencyName.getText();
      if (AgencyNameText == '代理店名で検索') {
        await RPA.WebBrowser.refresh();
      }
      // 更新画面をスルー
      try {
        const Koushin = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/header/div'}),5000);
        const KoushinText = await Koushin.getText();
        if (KoushinText.length > 1) {
          const NextButton01 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/footer/div[1]');
          await RPA.WebBrowser.mouseClick(NextButton01);
          await RPA.sleep(3000);
        }
      } catch {
        RPA.Logger.info('更新画面は出ませんでした');
      }
      await RPA.sleep(2000);
      const CampaignId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/table/tbody/tr[1]/td[3]');
      const CampaignIdText = await CampaignId.getText();
      if (CampaignIdText == JudgeCampaignId[0][0]) {
        LoadingFlag[0] = true;
        RPA.Logger.info('取得したキャンペーンID　　      　 →　' + CampaignIdText);
        RPA.Logger.info('貼り付けシートのキャンペーンID　   →　' + JudgeCampaignId);
        RPA.Logger.info('キャンペーンIDが一致しましたので、広告作成を開始します');
      }
    } catch {
      break;
    }
  }
  const CampaignName = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/table/tbody/tr[1]/td[3]');
  await RPA.WebBrowser.mouseClick(CampaignName);
  await RPA.sleep(3000);
  // 右下の広告作成を押下
  const CreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div/div/div[2]/div[3]/div/div[2]/div[2]/div'}),8000);
  await RPA.WebBrowser.mouseClick(CreateButton);
  await RPA.sleep(3000);
  // クリエイティブID(U列)を入力
  const CreativeId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[1]/div[2]/div/input');
  // クリエイティブIDがない場合、作業をスキップしてスタートに戻る
  const error = [['エラー']];
  if (WorkData[0][0][20].length < 1) {
    const ErrorText = [['シートにIDの記載がありません']];
    RPA.Logger.info(`エラー【${ErrorText}】` + ' 作業をスキップします');
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:ErrorText});
    return Start();
  } else {
    await RPA.WebBrowser.sendKeys(CreativeId,[WorkData[0][0][20]]);
    await RPA.sleep(5000);
    // エラーが起きた場合、作業をスキップしてスタートに戻る
    try {
      var ApplyFlag = ['false'];
      await RPA.sleep(300);
      const NotCreativeId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/div'}),1000);
      const NotCreativeIdText = await NotCreativeId.getText();
      if (String(NotCreativeIdText) == '該当する項目が存在しません') {
        RPA.Logger.info(`エラー【${NotCreativeIdText}】` + ' 作業をスキップします');
        const ErrorText2 = [['該当する項目が存在しません']];
        await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
        await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:ErrorText2});
        return Start();
      }
    } catch {
      ApplyFlag[0] = 'true';
      RPA.Logger.info('次の処理に進みます');
    }
    // クリエイティブ選択の「選択」をクリック
    const SelectCreative = await RPA.WebBrowser.findElementByXPath('//*[@id="reactroot"]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[1]/div');
    await RPA.WebBrowser.mouseClick(SelectCreative);
    // 配信期間をクリック
    const DateRange = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[2]/div[2]/div/div/div[1]/div');
    await RPA.WebBrowser.mouseClick(DateRange);
    // 配信期間：開始(C列)を入力
    const DateRangeStart = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[1]/div[1]/input');
    await DateRangeStart.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeStart,[WorkData[0][0][2]]);
    await RPA.sleep(300);
    // 配信期間：終了(D列)を入力
    const DateRangeEnd = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[2]/div[1]/input');
    await DateRangeEnd.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeEnd,[WorkData[0][0][3]]);
    await RPA.sleep(300);
    // 配信期間のOKボタンを押下
    const OKButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[3]/div/button[1]');
    await RPA.WebBrowser.mouseClick(OKButton);
    await RPA.sleep(1000);
    // 一度入力されている開始・終了期間をそれぞれ取得
    const DateRangeText = await DateRange.getText();
    RPA.Logger.info('AAAMSの開始日  → ' + DateRangeText);
    RPA.Logger.info('シートの開始日 → ' + WorkData[0][0][2] + ':00');
    const DateRange2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div');
    const DateRangeText2 = await DateRange2.getText();
    RPA.Logger.info('AAAMSの終了日  → ' + DateRangeText2);
    RPA.Logger.info('シートの終了日 → ' + WorkData[0][0][3] + ':00');
    if (DateRangeText == WorkData[0][0][2] + ':00' && DateRangeText2 == WorkData[0][0][3] + ':00') {
      RPA.Logger.info('期間が一致のため次に進みます');
    } else if (DateRangeText != WorkData[0][0][2] + ':00' && DateRangeText2 == WorkData[0][0][3] + ':00'){
      const NotMotchDateRangeStart = [['開始日が不一致です']];
      RPA.Logger.info(NotMotchDateRangeStart[0][0]);
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:NotMotchDateRangeStart});
      return Start();
    } else if (DateRangeText == WorkData[0][0][2] + ':00' && DateRangeText2 != WorkData[0][0][3] + ':00'){
      const NotMotchDateRangeEnd = [['終了日が不一致です']];
      RPA.Logger.info(NotMotchDateRangeEnd[0][0]);
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:NotMotchDateRangeEnd});
      return Start();
    } else if (DateRangeText != WorkData[0][0][2] + ':00' && DateRangeText2 != WorkData[0][0][3] + ':00'){
      const NotMotchDateRange = [['開始日・終了日が不一致です']];
      RPA.Logger.info(NotMotchDateRange[0][0]);
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!BA${Row[0]}:BA${Row[0]}`,values:NotMotchDateRange});
      return Start();
    }
    // imp比率(AA列)を入力
    if (WorkData[0][0][26].length < 1) {
      const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
      await ImpWeight.clear();
    } else {
      const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
      await ImpWeight.clear();
      await RPA.WebBrowser.sendKeys(ImpWeight,[WorkData[0][0][26]]);
    }
    // テキスト(AB列)を入力
    const Text = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
    if (WorkData[0][0][27] == '広告なしで ビデオをみる') {
      const WithoutAd = WorkData[0][0][27].slice(0, -6);
      await RPA.WebBrowser.sendKeys(Text,[WithoutAd]);
      Text.sendKeys(Key.ENTER);
    } else {
      await RPA.WebBrowser.sendKeys(Text,[WorkData[0][0][27]]);
      Text.sendKeys(Key.ENTER);
    }
    // URL(AC列)を入力
    const Url = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[2]/div[2]/div/input');
    if (WorkData[0][0][28].length < 1) {
      await RPA.WebBrowser.mouseClick(Url);
    } else {
      await RPA.WebBrowser.sendKeys(Url,[WorkData[0][0][28]]);
      await RPA.sleep(3000);
      await Url.sendKeys(Key.SPACE);
      await RPA.sleep(3000);
      await Url.sendKeys(Key.BACK_SPACE);
    }
    await RPA.sleep(3000);
    // URL(AC列)のエラーを判定するため、申込目標impをクリック
    const ImpRequest = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[7]/div[2]/div/input');
    await RPA.WebBrowser.mouseClick(ImpRequest);
    await RPA.sleep(3000);
    // エラーが出た場合は再入力
    for (var i = 0; i < 10; i++) {
      try {
        const FraudulentUrl = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[2]/div[1]/span'}),1000);
        const FraudulentUrlText = await FraudulentUrl.getText();
        RPA.Logger.info(FraudulentUrlText);
        if (String(FraudulentUrlText) == 'リンク先が不正です') {
          RPA.Logger.info('エラー文が出ましたので再入力します');
          await Url.clear();
          await RPA.sleep(100);
          await RPA.WebBrowser.sendKeys(Url,[WorkData[0][0][28]]);
          await RPA.sleep(3000);
          await Url.sendKeys(Key.SPACE);
          await RPA.sleep(3000);
          await Url.sendKeys(Key.BACK_SPACE);
          await RPA.sleep(3000);
          await RPA.WebBrowser.mouseClick(ImpRequest);
          await RPA.sleep(3000);
        } else {
          break;
        }
      } catch {
        RPA.Logger.info('エラー文が出ませんでしたので次に進みます');
        break;
      }
    }
  }
}


async function AdJudgeError2(WorkData, Row) {
  RPA.Logger.info('広告作成 エラー判定開始...');
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  if (WorkData[0][0][2].length < 1 || WorkData[0][0][3].length < 1 || WorkData[0][0][2].length < 1 && WorkData[0][0][3].length < 1) {
    const ErrorText3 = [['【広告作成】必須項目です']];
    RPA.Logger.info(`エラー【${ErrorText3}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AJ${Row[0]}:AJ${Row[0]}`,values:ErrorText3});
  }
  if (WorkData[0][0][26].length < 1) {
    const ErrorText4 = [['【広告作成】必須入力です']];
    RPA.Logger.info(`エラー【${ErrorText4}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AS${Row[0]}:AS${Row[0]}`,values:ErrorText4});
  }
  if (WorkData[0][0][27].length < 1) {
    const ErrorText5 = [['テキストを選択してください']];
    RPA.Logger.info(`エラー【${ErrorText5}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AX${Row[0]}:AX${Row[0]}`,values:ErrorText5});
  }
  if (WorkData[0][0][28].length < 1) {
    const ErrorText6 = [['URLを入力してください']];
    RPA.Logger.info(`エラー【${ErrorText6}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
  const NotUrl = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[6]/div[2]/div/div[2]/div[1]'}),1000);
  const NotUrlText = await NotUrl.getText();
  if (String(NotUrlText) == 'URLURLを入力してください') {
    RPA.Logger.info(`エラー【${NotUrlText}】`);
    const ErrorText6 = [['URLを入力してください']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
  if (String(NotUrlText) == 'URLURLを正しく入力してください') {
    RPA.Logger.info(`エラー【${NotUrlText}】`);
    const ErrorText6 = [['URLを正しく入力してください']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
  if (String(NotUrlText) == 'URLリンク先が不正です') {
    RPA.Logger.info(`エラー【${NotUrlText}】`);
    const ErrorText6 = [['リンク先が不正です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AY${Row[0]}:AY${Row[0]}`,values:ErrorText6});
  }
}


async function GetAdvertisementId2(Row) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    return Start();
  }
  // OKボタンをクリック
  const OKButton2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton2);
  RPA.Logger.info('広告 ID発番中...');
  await RPA.sleep(5000);
  // 発番した広告IDの最右側の「・・・」をマウスオーバー
  const BalloonMenu = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[10]/div'
      }),
      5000
    );
  await RPA.WebBrowser.mouseMove(BalloonMenu);
  await RPA.sleep(1000);
  // 「配信を変更する」をクリック
  const ChangeDelivery = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[3]/div/div[4]');
  await RPA.WebBrowser.mouseClick(ChangeDelivery);
  await RPA.sleep(1000);
  const OKButton3 = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton3);
  await RPA.sleep(8000);
  const GetAdvertisementId = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[1]');
  const AdvertisementId = await GetAdvertisementId.getText();
  RPA.Logger.info(AdvertisementId);
  // キャンペーン広告貼り付け（外部リンク付き）シートのAW列に発番した広告IDを記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName4}!AW${Row[0] + 1}:AW${Row[0] + 1}`,values:[[AdvertisementId]]});
  RPA.Logger.info('広告 作成完了しました');
  // 作業完了を記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName3}!AG${Row[0]}:AG${Row[0]}`,values:[['完了']]});
}