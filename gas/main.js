// --- グローバル定数定義 ---
const CURRENT_PARTICIPANTS_SHEET_NAME = "今月の参加者";
const PAST_TABLE_DATA_SHEET_NAME = "過去の参加者";
const SETTINGS_SHEET_NAME = "設定";
const OUTPUT_SHEET_NAME = "今月の席割";

const CHECKBOX_GENERATE_ASSIGNMENT_CELL = "A3"; // 席割り実行トリガーのチェックボックスセル
const CHECKBOX_REGISTER_PAST_DATA_CELL = "D3"; // 過去データ登録トリガーのチェックボックスセル

const OUTPUT_START_CELL = "A6"; // 席割り結果の出力開始セル
const EVALUATION_START_CELL = "L6"; // 評価コメントの出力開始セル

// 「今月の参加者」シートの列インデックス (0-indexed)
const COL_NAME = 0;
const COL_VENUE = 1;
const COL_MEMBERSHIP = 2;
const COL_TABLE_1ST = 3;
const COL_TABLE_2ND = 4;
const COL_TABLE_3RD = 5;
const COL_MANAGEMENT = 6;
const COL_CARETAKER = 7;
const COL_LEADER = 8; // テーブルリーダー列
const COL_INTRODUCER = 9;

// 「過去の参加者」シートの列インデックス (0-indexed)
// 実際のスプレッドシートのヘッダーの並び順に合わせる
const PAST_COL_HOLDING_MONTH = 0; // 開催月
const PAST_COL_TABLE_NUM = 1; // 卓番
const PAST_COL_NAME = 2; // 名前
const PAST_COL_VENUE = 3; // 所属会場

// --- ユーティリティ関数（データ取得・設定関連） ---

/**
 * 「設定」シートから設定値を取得する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 「設定」シートオブジェクト
 * @returns {Object} 設定値オブジェクト
 */
function getSettings(sheet) {
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "自会場") settings.selfVenueName = data[i][1];
    if (data[i][0] === "車座回数") settings.rotationCount = data[i][1];
    if (data[i][0] === "テーブル数") settings.tableCount = data[i][1];
    if (data[i][0] === "今月") settings.currentMonth = data[i][1]; // 「今月」設定も取得
  }
  return settings;
}

/**
 * 「過去の参加者」シートから生のデータを取得する（会場名補完用）。
 * ヘッダーのインデックス検索を堅牢化。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 「過去の参加者」シートオブジェクト
 * @returns {Array<Object>} 過去の記録オブジェクトの配列
 */
function getPastTableDataRaw(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const nameCol = headers.indexOf("名前");
  const venueCol = headers.indexOf("所属会場");
  const holdingMonthCol = headers.indexOf("開催月");

  if (nameCol === -1 || venueCol === -1 || holdingMonthCol === -1) {
    Logger.log(
      "getPastTableDataRaw: 過去の参加者シートの必須ヘッダーが見つかりません。名前, 所属会場, 開催月"
    );
    return [];
  }

  const pastRecords = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    pastRecords.push({
      name: String(row[nameCol]).trim(),
      venue: String(row[venueCol]).trim(),
      holdingMonth: row[holdingMonthCol],
    });
  }
  return pastRecords;
}

/**
 * 「今月の参加者」シートからデータを取得し、会場名を補完する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 「今月の参加者」シートオブジェクト
 * @param {GoogleAppsScript.Spreadsheet.Sheet} pastDataSheet - 「過去の参加者」シートオブジェクト（会場名補完用）
 * @param {string} selfVenueName - 自会場名
 * @returns {Array<Object>} 参加者オブジェクトの配列
 */
function getCurrentParticipantsData(sheet, pastDataSheet, selfVenueName) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Browser.msgBox(
      "データ不足",
      "「今月の参加者」シートにヘッダー行以外のデータがありません。",
      Browser.Buttons.OK
    );
    return [];
  }

  const participants = [];
  const pastTableDataRaw = getPastTableDataRaw(pastDataSheet); // 会場名補完用

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const participant = {};

    // 定義された列インデックスに基づいてデータをマッピング (空白も考慮)
    participant.name =
      row[COL_NAME] !== undefined ? String(row[COL_NAME]).trim() : "";
    participant.venue =
      row[COL_VENUE] !== undefined ? String(row[COL_VENUE]).trim() : "";
    participant.membership =
      row[COL_MEMBERSHIP] !== undefined
        ? String(row[COL_MEMBERSHIP]).trim()
        : "";
    participant.table1st =
      row[COL_TABLE_1ST] !== undefined &&
      String(row[COL_TABLE_1ST]).trim() !== ""
        ? String(row[COL_TABLE_1ST]).trim().toUpperCase()
        : null;
    participant.table2nd =
      row[COL_TABLE_2ND] !== undefined &&
      String(row[COL_TABLE_2ND]).trim() !== ""
        ? String(row[COL_TABLE_2ND]).trim().toUpperCase()
        : null;
    participant.table3rd =
      row[COL_TABLE_3RD] !== undefined &&
      String(row[COL_TABLE_3RD]).trim() !== ""
        ? String(row[COL_TABLE_3RD]).trim().toUpperCase()
        : null;
    participant.management = row[COL_MANAGEMENT] == 1;
    participant.caretaker = row[COL_CARETAKER] == 1;
    participant.leader = row[COL_LEADER] == 1;
    participant.introducer =
      row[COL_INTRODUCER] !== undefined &&
      String(row[COL_INTRODUCER]).trim() !== ""
        ? String(row[COL_INTRODUCER]).trim()
        : null;

    // 必須項目が空の場合はスキップし、警告
    if (!participant.name) {
      Browser.msgBox(
        "データエラー",
        `「今月の参加者」シートの${
          i + 1
        }行目に名前がありません。この行はスキップされます。`,
        Browser.Buttons.OK
      );
      continue;
    }
    if (!participant.membership) {
      Browser.msgBox(
        "データエラー",
        `「今月の参加者」シートの${i + 1}行目（${
          participant.name
        }）に会員区分がありません。この行はスキップされます。`,
        Browser.Buttons.OK
      );
      continue;
    }

    // 会場名が空の場合、過去のデータから補完。過去にもなければ自会場とする
    if (!participant.venue) {
      const lastVenueRecord = pastTableDataRaw.find(
        (p) => p.name === participant.name && p.venue
      );
      if (lastVenueRecord) {
        participant.venue = lastVenueRecord.venue;
      } else {
        participant.venue = selfVenueName;
      }
    }
    participants.push(participant);
  }
  return participants;
}

/**
 * 「過去の参加者」シートから席割りロジック用のデータを取得する（過去6ヶ月分）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 「過去の参加者」シートオブジェクト
 * @returns {Array<Object>} 過去の記録オブジェクトの配列
 */
function getPastTableData(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // ヘッダーのみの場合

  const headers = data[0];
  const nameCol = headers.indexOf("名前");
  const tableNumCol = headers.indexOf("卓番");
  const holdingMonthCol = headers.indexOf("開催月");
  const venueCol = headers.indexOf("所属会場");

  if (
    nameCol === -1 ||
    tableNumCol === -1 ||
    holdingMonthCol === -1 ||
    venueCol === -1
  ) {
    Logger.log(
      "getPastTableData: 過去の参加者シートの必須ヘッダーが見つかりません。名前, 卓番, 開催月, 所属会場"
    );
    Browser.msgBox(
      "データエラー",
      "「過去の参加者」シートのヘッダーが不完全です。席割り精度に影響します。",
      Browser.Buttons.OK
    );
    return [];
  }

  const settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  const settings = getSettings(settingsSheet);
  const currentMonthNum = settings.currentMonth;

  if (!currentMonthNum || typeof currentMonthNum !== "number") {
    Browser.msgBox(
      "エラー",
      "「設定」シートの「今月」がYYYYMM形式の数値でありません。過去データ参照ができません。",
      Browser.Buttons.OK
    );
    return [];
  }

  const pastData = [];
  const currentYear = Math.floor(currentMonthNum / 100);
  const currentMonth = currentMonthNum % 100;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pastHoldingMonthNum = row[holdingMonthCol];

    if (
      typeof pastHoldingMonthNum !== "number" ||
      String(pastHoldingMonthNum).length !== 6
    ) {
      Logger.log(
        `Skipping past data row ${
          i + 1
        } due to invalid HoldingMonth: ${pastHoldingMonthNum}`
      );
      continue;
    }

    const pastYear = Math.floor(pastHoldingMonthNum / 100);
    const pastMonth = pastHoldingMonthNum % 100;

    // 月差を計算 (現在の月を0ヶ月前として、1～6ヶ月前のデータを取得)
    const monthDiff =
      (currentYear - pastYear) * 12 + (currentMonth - pastMonth);

    if (monthDiff >= 1 && monthDiff <= 6) {
      // 過去6ヶ月以内のデータのみを抽出
      pastData.push({
        name: String(row[nameCol]).trim(),
        tableNum: String(row[tableNumCol]).trim().toUpperCase(),
        period: monthDiff, // 月差をperiodとして使用
        venue: String(row[venueCol]).trim(),
      });
    }
  }
  return pastData;
}

// --- シート初期設定関数（プルダウン設定用） ---

/**
 * スプレッドシートが開かれたときに実行される関数。
 * カスタムメニューを追加し、プルダウン設定を促す。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("席割りシステム")
    .addItem(
      "今月の参加者シートにプルダウンを設定",
      "setupParticipantsSheetDropdowns"
    )
    .addToUi();
}

/**
 * 「今月の参加者」シートの卓番列と会員区分列にプルダウン（データ検証）を設定する。
 * この関数は、Apps Scriptエディタのメニューから手動で実行されることを想定。
 */
function setupParticipantsSheetDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName(CURRENT_PARTICIPANTS_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!participantsSheet || !settingsSheet) {
    Browser.msgBox(
      "エラー",
      "「今月の参加者」または「設定」シートが見つかりません。シート名を確認してください。",
      Browser.Buttons.OK
    );
    return;
  }

  const settings = getSettings(settingsSheet);
  const tableCount = settings.tableCount;

  if (!tableCount || typeof tableCount !== "number" || tableCount <= 0) {
    Browser.msgBox(
      "エラー",
      "「設定」シートの「テーブル数」が正しく設定されていません。1以上の数値を入力してください。",
      Browser.Buttons.OK
    );
    return;
  }

  // 卓番のプルダウンリストを生成 (A, B, C...)
  const tableNumberList = [];
  for (let i = 0; i < tableCount; i++) {
    tableNumberList.push(String.fromCharCode(65 + i)); // ASCII A=65
  }

  // 会員区分のプルダウンリスト
  const membershipList = ["正会員以上", "準会員", "ゲスト"];

  // データ検証ルールの適用範囲をシートの最大行までとする
  const numRowsToApply = participantsSheet.getMaxRows() - 1;

  // 卓番のプルダウン (1順目卓番, 2順目卓番, 3順目卓番)
  const tableNumberRanges = [
    participantsSheet.getRange(2, COL_TABLE_1ST + 1, numRowsToApply), // +1 は1-indexedへの変換
    participantsSheet.getRange(2, COL_TABLE_2ND + 1, numRowsToApply),
    participantsSheet.getRange(2, COL_TABLE_3RD + 1, numRowsToApply),
  ];

  tableNumberRanges.forEach((range) => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(tableNumberList)
      .setAllowInvalid(true) // リストにない値も許可（ただし警告表示）
      .setHelpText("設定シートのテーブル数に応じた卓番を選択してください。")
      .build();
    range.setDataValidation(rule);
  });

  // 会員区分のプルダウン
  const membershipRange = participantsSheet.getRange(
    2,
    COL_MEMBERSHIP + 1,
    numRowsToApply
  );
  const membershipRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(membershipList)
    .setAllowInvalid(false) // リストにない値は拒否
    .setHelpText(
      "「正会員以上」「準会員」「ゲスト」のいずれかを選択してください。"
    )
    .build();
  membershipRange.setDataValidation(membershipRule);

  Browser.msgBox(
    "プルダウン設定完了",
    "「今月の参加者」シートの卓番列と会員区分列にプルダウンが設定されました。",
    Browser.Buttons.OK
  );
}

// --- メイン実行トリガー関数 ---

/**
 * 「今月の席割」シートのA3セルまたはD3セルが変更されたときにトリガーされる関数。
 * 席割り生成または過去データ登録のメインプロセスを開始する。
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // A3チェックボックスがONになったら席割り生成
  if (
    sheet.getName() === OUTPUT_SHEET_NAME &&
    range.getA1Notation() === CHECKBOX_GENERATE_ASSIGNMENT_CELL &&
    range.getValue() === true
  ) {
    generateAndEvaluateTableAssignment();
  }
  // D3チェックボックスがONになったら過去データ登録
  if (
    sheet.getName() === OUTPUT_SHEET_NAME &&
    range.getA1Notation() === CHECKBOX_REGISTER_PAST_DATA_CELL &&
    range.getValue() === true
  ) {
    registerCurrentAssignmentToPastData();
  }
}

/**
 * 席割り生成と評価の全体プロセスを管理する関数。
 */
function generateAndEvaluateTableAssignment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentParticipantsSheet = ss.getSheetByName(
    CURRENT_PARTICIPANTS_SHEET_NAME
  );
  const pastTableDataSheet = ss.getSheetByName(PAST_TABLE_DATA_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const outputSheet = ss.getSheetByName(OUTPUT_SHEET_NAME);

  // 必要なシートが存在するかチェック
  if (
    !currentParticipantsSheet ||
    !pastTableDataSheet ||
    !settingsSheet ||
    !outputSheet
  ) {
    Browser.msgBox(
      "エラー",
      "必要なシートが見つかりません。シート名を確認してください。\n「今月の参加者」「過去の参加者」「設定」「今月の席割」",
      Browser.Buttons.OK
    );
    return;
  }

  // 既存の出力と評価をクリア
  outputSheet
    .getRange(OUTPUT_START_CELL + ":" + outputSheet.getMaxColumns())
    .clearContent();
  outputSheet
    .getRange(EVALUATION_START_CELL + ":" + outputSheet.getMaxColumns())
    .clearContent();

  const settings = getSettings(settingsSheet);

  // 設定値の検証
  if (
    !settings.rotationCount ||
    !settings.tableCount ||
    typeof settings.tableCount !== "number" ||
    settings.tableCount <= 0 ||
    !settings.selfVenueName ||
    !settings.currentMonth
  ) {
    Browser.msgBox(
      "エラー",
      "「設定」シートに必須項目が正しく設定されていません。「自会場」「車座回数」「テーブル数」「今月」をすべて入力してください。",
      Browser.Buttons.OK
    );
    return;
  }
  if (
    typeof settings.currentMonth !== "number" ||
    String(settings.currentMonth).length !== 6
  ) {
    Browser.msgBox(
      "エラー",
      "「設定」シートの「今月」がYYYYMM形式の数値でありません。例: 202506。",
      Browser.Buttons.OK
    );
    return;
  }

  const participantsData = getCurrentParticipantsData(
    currentParticipantsSheet,
    pastTableDataSheet,
    settings.selfVenueName
  );
  const pastTableData = getPastTableData(pastTableDataSheet); // 過去6ヶ月分のデータのみ

  if (participantsData.length === 0) {
    Browser.msgBox(
      "エラー",
      "「今月の参加者」シートにデータがありません。席割りを行うには参加者データを入力してください。",
      Browser.Buttons.OK
    );
    return;
  }

  // 席割りロジックの実行
  const result = assignTables(participantsData, pastTableData, settings);

  // assignTablesからエラーが返された場合は処理を中断
  if (result.error) {
    Browser.msgBox("席割りエラー", result.errorMessage, Browser.Buttons.OK);
    outputSheet
      .getRange(EVALUATION_START_CELL)
      .setValue("席割り処理が中断されました。" + result.errorMessage);
    return;
  }

  // 席割りの出力
  outputTableAssignment(outputSheet, result.tableAssignment);

  // 評価の出力
  outputEvaluation(outputSheet, result.evaluation);
}

// --- 席割りロジック関連関数 ---

/**
 * 席割りロジックのメイン関数
 * @param {Array<Object>} participants - 今月の参加者データ（全参加者）
 * @param {Array<Object>} pastData - 過去のテーブルデータ（過去6ヶ月分）
 * @param {Object} settings - 設定データ (selfVenueName, rotationCount, tableCount)
 * @returns {Object} - 生成されたテーブル割り と 評価。エラー時は {error: true, errorMessage: string} を返す
 */
function assignTables(participants, pastData, settings) {
  let finalTableAssignment = [];
  let evaluation = [];
  const rotationCount = settings.rotationCount;
  const tableCount = settings.tableCount;
  const selfVenueName = settings.selfVenueName;
  const totalParticipants = participants.length;

  // 各テーブルの目標人数範囲
  const targetMinPerTable = Math.floor(totalParticipants / tableCount);
  const targetMaxPerTable =
    targetMinPerTable + (totalParticipants % tableCount > 0 ? 1 : 0);

  // 全巡目で固定されるテーブルリーダーを事前に処理
  const allLeaders = participants.filter((p) => p.leader);
  const fixedLeaderAssignments = {}; // {leaderName: fixedTableNum}

  // リーダーの固定卓番を決定し、ルール違反があれば処理中断
  for (const leader of allLeaders) {
    let fixedTableNum = leader.table1st;
    if (!fixedTableNum) {
      return {
        error: true,
        errorMessage: `テーブルリーダー「${leader.name}」には1巡目の卓番指定がありません。席割り処理を中断します。`,
      };
    }
    if (leader.table2nd && leader.table2nd !== fixedTableNum) {
      evaluation.push(
        `警告: テーブルリーダー「${leader.name}」の2巡目卓番(${leader.table2nd})が1巡目卓番(${fixedTableNum})と異なります。1巡目の卓番に固定されます。`
      );
    }
    if (leader.table3rd && leader.table3rd !== fixedTableNum) {
      evaluation.push(
        `警告: テーブルリーダー「${leader.name}」の3巡目卓番(${leader.table3rd})が1巡目卓番(${fixedTableNum})と異なります。1巡目の卓番に固定されます。`
      );
    }
    fixedLeaderAssignments[leader.name] = fixedTableNum;
  }

  // 各巡目ごとのテーブル割りを作成
  for (let r = 1; r <= rotationCount; r++) {
    evaluation.push(`## ${r}巡目 席割り評価`);

    // この巡でまだ割り当てられていない参加者のリスト (毎回初期化)
    let currentParticipantsPool = JSON.parse(JSON.stringify(participants));
    // この巡で既に割り当て済みの参加者の名前を追跡するセット (毎回初期化)
    let assignedNamesInRotation = new Set();

    // ヘルパー関数：メンバーをテーブルに割り当て、追跡セットを更新
    // 割り当てが成功したらtrueを返す。重複割り当ては行わない。
    const assignMemberToTable = (member, table, isManual = false) => {
      if (assignedNamesInRotation.has(member.name)) {
        return false; // 既に割り当て済みであれば何もしない
      }
      table.members.push(member);
      if (isManual) {
        table.manualMembers.push(member);
      } else {
        table.autoMembers.push(member);
      }
      if (member.management) table.managementCount++;
      if (member.caretaker) table.caretakerCount++;
      if (member.leader) table.leaderCount++;
      if (member.membership === "ゲスト") table.guestCount++;
      if (member.venue !== selfVenueName) {
        table.otherVenueCounts[member.venue] =
          (table.otherVenueCounts[member.venue] || 0) + 1;
      }
      assignedNamesInRotation.add(member.name);
      return true;
    };

    // テーブルの初期化
    let tables = [];
    for (let i = 0; i < tableCount; i++) {
      tables.push({
        name: String.fromCharCode(65 + i), // A, B, C...
        members: [],
        manualMembers: [],
        autoMembers: [],
        managementCount: 0,
        caretakerCount: 0,
        otherVenueCounts: {},
        leaderCount: 0,
        guestCount: 0,
      });
    }

    // --- Phase 1: 固定卓番のテーブルリーダーを割り当て（最優先） ---
    // currentParticipantsPool からリーダーを抽出し、処理。割り当て済みのリーダーはプールから除外
    const leadersToProcess = currentParticipantsPool.filter(
      (p) => p.leader && fixedLeaderAssignments[p.name]
    );

    for (const leader of leadersToProcess) {
      const fixedTableNum = fixedLeaderAssignments[leader.name];
      const targetTable = tables.find((t) => t.name === fixedTableNum);

      // テーブルリーダーの「1名のみ」ルールと、卓番の存在を厳格にチェックし、違反なら中断
      if (!targetTable) {
        return {
          error: true,
          errorMessage: `テーブルリーダー「${leader.name}」（${r}巡目）の指定卓番「${fixedTableNum}」が見つかりません。「設定」シートのテーブル数範囲外か確認してください。席割り処理を中断します。`,
        };
      }
      if (targetTable.leaderCount >= 1) {
        // 既にリーダーがいる
        return {
          error: true,
          errorMessage: `テーブルリーダー「${leader.name}」（${r}巡目）の指定卓番「${fixedTableNum}」には既に別のテーブルリーダーがいます。各テーブルにリーダーは1名のみです。席割り処理を中断します。`,
        };
      }

      assignMemberToTable(leader, targetTable, true);
    }
    // 割り当てられたリーダーを currentParticipantsPool から除外
    currentParticipantsPool = currentParticipantsPool.filter(
      (p) => !assignedNamesInRotation.has(p.name)
    );

    // --- Phase 2: その他の手動指定メンバーを割り当て ---
    // currentParticipantsPool から、このフェーズでマニュアル指定のあるメンバーを抽出
    const manualOtherMembersToProcess = currentParticipantsPool.filter((p) => {
      let targetTableNum = null;
      if (r === 1) targetTableNum = p.table1st;
      if (r === 2) targetTableNum = p.table2nd;
      if (r === 3) targetTableNum = p.table3rd;

      return targetTableNum && !p.leader;
    });

    for (const member of manualOtherMembersToProcess) {
      if (assignedNamesInRotation.has(member.name)) continue;

      let targetTableNum = null;
      if (r === 1) targetTableNum = member.table1st;
      if (r === 2) targetTableNum = member.table2nd;
      if (r === 3) targetTableNum = member.table3rd;

      const targetTable = tables.find((t) => t.name === targetTableNum);
      // マニュアル割り当ては人数上限を**超えても割り当てを試みる**が、評価で警告
      if (targetTable && assignMemberToTable(member, targetTable, true)) {
        if (targetTable.members.length > targetMaxPerTable) {
          evaluation.push(
            `警告: ${member.name} (${r}巡目) の指定卓番「${targetTableNum}」への割り当てにより、人数上限(${targetMaxPerTable}名)を超過しました（現在${targetTable.members.length}名）。`
          );
        }
      } else {
        evaluation.push(
          `警告: ${member.name} (${r}巡目) の指定卓番「${targetTableNum}」に割り当てられませんでした（理由：卓番が見つからない）。このメンバーは自動割り当てされます。`
        );
      }
    }
    // 割り当てられたメンバーを currentParticipantsPool から除外
    currentParticipantsPool = currentParticipantsPool.filter(
      (p) => !assignedNamesInRotation.has(p.name)
    );

    // --- Phase 3: ゲストと紹介者のグループ化 (自動割り当て対象のみ) ---
    let guestGroups = [];
    let tempParticipantsForGuestGrouping = [...currentParticipantsPool];
    let processedForGuestGrouping = new Set();

    for (const p of tempParticipantsForGuestGrouping) {
      if (assignedNamesInRotation.has(p.name)) continue;
      if (
        p.membership === "ゲスト" &&
        p.introducer &&
        !processedForGuestGrouping.has(p.name)
      ) {
        const introducer = currentParticipantsPool.find(
          (rp) => rp.name === p.introducer
        );
        if (introducer && !processedForGuestGrouping.has(introducer.name)) {
          const guestsForIntroducer = currentParticipantsPool.filter(
            (rp) =>
              rp.introducer === introducer.name &&
              rp.membership === "ゲスト" &&
              !processedForGuestGrouping.has(rp.name)
          );
          const group = [introducer, ...guestsForIntroducer];
          guestGroups.push(group);
          group.forEach((member) => processedForGuestGrouping.add(member.name));
        }
      }
    }
    // currentParticipantsPool からグループ化されたメンバーを除外
    currentParticipantsPool = currentParticipantsPool.filter(
      (p) => !processedForGuestGrouping.has(p.name)
    );

    // --- Phase 4: 残りの自動割り当て対象メンバーの割り当て（スコアリングベース） ---
    // 運営部、世話人、自動割り当て対象リーダー、他会場、自会場の順に候補を結合
    let allAutoAssignCandidates = [
      ...currentParticipantsPool.filter((p) => p.management),
      ...currentParticipantsPool.filter((p) => p.caretaker),
      ...currentParticipantsPool.filter((p) => p.leader),
      // 他会場メンバーは、運営部/世話人/リーダーではないもの
      ...currentParticipantsPool.filter(
        (p) =>
          p.venue !== selfVenueName &&
          !p.management &&
          !p.caretaker &&
          !p.leader
      ),
      // 自会場メンバーは、運営部/世話人/リーダーではないもの
      ...currentParticipantsPool.filter(
        (p) =>
          p.venue === selfVenueName &&
          !p.management &&
          !p.caretaker &&
          !p.leader
      ),
    ];
    allAutoAssignCandidates.sort(() => Math.random() - 0.5); // ランダムにシャッフル

    // まずゲストグループを割り当てる (スコアリング対象外だが、人数と他会場考慮)
    for (const group of guestGroups) {
      if (group.some((member) => assignedNamesInRotation.has(member.name)))
        continue; // グループの誰か一人でも割り当て済みならスキップ

      let assigned = false;
      // 割り当て可能なテーブルをフィルタリング
      let validTablesForGroup = tables.filter((table) => {
        // 他会場同席ルールチェック
        for (const member of group) {
          if (
            member.venue !== selfVenueName &&
            table.otherVenueCounts[member.venue]
          ) {
            return false; // ルール違反
          }
        }
        // 人数上限チェック
        return table.members.length + group.length <= targetMaxPerTable;
      });

      if (validTablesForGroup.length > 0) {
        validTablesForGroup.sort((a, b) => a.members.length - b.members.length); // 最も人数が少ないテーブルを優先
        group.forEach((member) =>
          assignMemberToTable(member, validTablesForGroup[0], false)
        );
        assigned = true;
      }

      if (!assigned) {
        evaluation.push(
          `警告: ${r}巡目: ゲストグループ（紹介者:${group[0].name}）をどのテーブルにも割り当てられませんでした（人数上限または他会場ルールのため）。強制割り当て。`
        );
        tables.sort((a, b) => a.members.length - b.members.length); // 最も人数が少ないテーブルに強制割り当て
        group.forEach((member) =>
          assignMemberToTable(member, tables[0], false)
        );
      }
    }

    // その他の自動割り当てメンバーの割り当て（スコアリングで最適化）
    for (const participant of allAutoAssignCandidates) {
      if (assignedNamesInRotation.has(participant.name)) continue; // 既に割り当て済みならスキップ

      let bestTable = null;
      let highestScore = -Infinity;
      let availableTables = [];

      for (let i = 0; i < tables.length; i++) {
        const table = tables[i];

        // isAssignmentValid関数で絶対ルールをチェック (パフォーマンスと正確性のため)
        // isAssignmentValidで割り当て済み、他会場同席、人数上限、テーブルリーダー重複をチェックする
        const isValid = isAssignmentValid(
          participant,
          table,
          pastData,
          selfVenueName,
          totalParticipants,
          settings,
          assignedNamesInRotation
        );
        if (!isValid) continue; // 絶対ルールに違反したらスキップ

        availableTables.push(table);
      }

      if (availableTables.length > 0) {
        // 割り当て可能なテーブルがある場合のみスコアリング
        for (const table of availableTables) {
          // 候補となるテーブルのクローンを作成して、仮割り当て後の状態をシミュレーション
          const tempTable = JSON.parse(JSON.stringify(table));
          tempTable.members.push(participant); // メンバーを仮追加
          // 仮追加後の属性カウントを更新
          if (participant.management) tempTable.managementCount++;
          if (participant.caretaker) tempTable.caretakerCount++;
          if (participant.leader) tempTable.leaderCount++;
          if (participant.venue !== selfVenueName) {
            tempTable.otherVenueCounts[participant.venue] =
              (tempTable.otherVenueCounts[participant.venue] || 0) + 1;
          }

          // calculateAssignmentScore に participants (全参加者) を渡す
          const score = calculateAssignmentScore(
            participant,
            tempTable,
            participants,
            pastData,
            selfVenueName,
            totalParticipants,
            settings
          );

          if (score > highestScore) {
            highestScore = score;
            bestTable = table;
          } else if (score === highestScore) {
            if (Math.random() < 0.5) {
              // 同点の場合、ランダム性を加える
              bestTable = table;
            }
          }
        }
      }

      if (bestTable) {
        assignMemberToTable(participant, bestTable, false);
      } else {
        // 全てのテーブルに割り当てられない（厳格なルールのため）場合
        evaluation.push(
          `警告: ${r}巡目: ${participant.name} をどのテーブルにも割り当てられませんでした（全ての厳格なルールを満たすテーブルがないため）。強制割り当て。`
        );
        tables.sort((a, b) => a.members.length - b.members.length); // 最も人数が少ないテーブルに強制割り当て
        assignMemberToTable(participant, tables[0], false);
      }
    }

    // --- 巡目ごとのルール適合性評価 ---
    tables.forEach((table) => {
      evaluation.push(`### ${r}巡目 テーブル${table.name}`);

      // 1. テーブルリーダーの確認
      const designatedLeadersInTable = table.members.filter((m) => m.leader);
      if (designatedLeadersInTable.length === 1) {
        evaluation.push(
          `- **テーブルリーダー**: リーダー指定されたメンバーが1名います。適合。`
        );
      } else if (designatedLeadersInTable.length > 1) {
        evaluation.push(
          `- **テーブルリーダー**: リーダー指定されたメンバーが複数名います（${designatedLeadersInTable.length}名）。ルール違反です。`
        );
      } else {
        evaluation.push(
          `- **テーブルリーダー**: リーダー指定されたメンバーが見当たりません。ルール違反です。`
        );
      }

      // 2. ゲストと紹介者、ゲスト人数
      let guestsInTable = table.members.filter(
        (m) => m.membership === "ゲスト"
      );
      let guestRuleViolations = [];
      let guestIntroducerGroups = {};

      guestsInTable.forEach((g) => {
        if (g.introducer) {
          if (!guestIntroducerGroups[g.introducer]) {
            guestIntroducerGroups[g.introducer] = [];
          }
          guestIntroducerGroups[g.introducer].push(g);
          const introducerFound = table.members.some(
            (m) => m.name === g.introducer
          );
          if (!introducerFound) {
            guestRuleViolations.push(
              `ゲスト「${g.name}」の紹介者「${g.introducer}」がこのテーブルにいません。`
            );
          }
        } else {
          guestRuleViolations.push(
            `ゲスト「${g.name}」には紹介者が指定されていません。`
          );
        }
      });

      if (
        Object.keys(guestIntroducerGroups).length === 0 &&
        guestsInTable.length === 0
      ) {
        evaluation.push(`- **ゲスト配置**: このテーブルにはゲストがいません。`);
      } else if (guestRuleViolations.length > 0) {
        guestRuleViolations.forEach((v) =>
          evaluation.push(`- **ゲスト配置**: ${v}`)
        );
      } else {
        if (Object.keys(guestIntroducerGroups).length > 1) {
          evaluation.push(
            `- **ゲスト人数**: 異なる紹介者からのゲストが複数名います。極力1名に収めるルールに反しています。`
          );
        } else {
          evaluation.push(
            `- **ゲスト人数**: ゲストは${guestsInTable.length}名です。同一紹介者のゲストが複数いる場合は許容されます。`
          );
        }
        evaluation.push(
          `- **ゲスト配置**: ゲストの配置はルールに適合しています。`
        );
      }

      // 3. Member Count
      if (
        table.members.length < targetMinPerTable ||
        table.members.length > targetMaxPerTable
      ) {
        evaluation.push(
          `- **人数**: 人数が${table.members.length}名で、目標人数（${targetMinPerTable}-${targetMaxPerTable}名）の範囲外です。`
        );
      } else {
        evaluation.push(
          `- **人数**: 人数が${table.members.length}名で、目標人数範囲内です。`
        );
      }

      // 4. Same Other-Venue Conflict
      let venueConflicts = [];
      for (const venue in table.otherVenueCounts) {
        if (table.otherVenueCounts[venue] > 1) {
          venueConflicts.push(
            `${venue}会場 (${table.otherVenueCounts[venue]}名)`
          );
        }
      }
      if (venueConflicts.length > 0) {
        evaluation.push(
          `- **他会場の同席**: 同じ他会場の方（${venueConflicts.join(
            ", "
          )}）が同席しています。`
        );
      } else {
        evaluation.push(
          `- **他会場の同席**: 同じ他会場の方同士の同席はありません。`
        );
      }

      // 5. Other-Venue and Self-Venue Balance
      const selfVenueCount = table.members.filter(
        (m) => m.venue === selfVenueName
      ).length;
      const otherVenueTotalCount = table.members.length - selfVenueCount;
      evaluation.push(
        `- **会場バランス**: 自会場 ${selfVenueCount}名、他会場 ${otherVenueTotalCount}名。`
      );

      // 6. Past Same-Seat Avoidance (Auto-assigned members only)
      let pastSameSeatViolations = [];
      const autoMembersOnly = table.autoMembers;
      for (let i = 0; i < autoMembersOnly.length; i++) {
        for (let j = i + 1; j < autoMembersOnly.length; j++) {
          const member1 = autoMembersOnly[i];
          const member2 = autoMembersOnly[j];
          const hasSatTogetherRecently = pastData.some(
            (past) =>
              past.name === member1.name &&
              pastData.some(
                (past2) =>
                  past2.name === member2.name &&
                  past.tableNum === past2.tableNum &&
                  past.period <= 6
              )
          );
          if (hasSatTogetherRecently) {
            pastSameSeatViolations.push(
              `「${member1.name}」と「${member2.name}」`
            );
          }
        }
      }
      if (pastSameSeatViolations.length > 0) {
        evaluation.push(
          `- **過去の同席回避**: 過去6ヶ月以内に同席した可能性のある自動割り当てペアがいます: ${pastSameSeatViolations.join(
            ", "
          )}。`
        );
      } else {
        evaluation.push(
          `- **過去の同席回避**: 過去6ヶ月以内に同席した自動割り当てペアはいません。`
        );
      }
    });

    // Overall evaluation for Management distribution
    const totalManagementMembers = participants.filter(
      (p) => p.management
    ).length;
    let managementDistribution = tables.map((table) => table.managementCount);
    let isManagementEven = true;
    if (tables.length > 0 && totalManagementMembers > 0) {
      const minMgmt = Math.min(...managementDistribution);
      const maxMgmt = Math.max(...managementDistribution);
      if (maxMgmt - minMgmt > 1) {
        isManagementEven = false;
      }
    }
    if (isManagementEven) {
      evaluation.push(
        `- **運営部配置**: 運営部の配置は比較的均等です (${managementDistribution.join(
          ", "
        )})。`
      );
    } else {
      evaluation.push(
        `- **運営部配置**: 運営部の配置は均等ではありません (${managementDistribution.join(
          ", "
        )})。`
      );
    }

    // Overall evaluation for Caretaker distribution
    const totalCaretakerMembers = participants.filter(
      (p) => p.caretaker
    ).length;
    let caretakerDistribution = tables.map((table) => table.caretakerCount);
    let isCaretakerEven = true;
    if (tables.length > 0 && totalCaretakerMembers > 0) {
      const minCare = Math.min(...caretakerDistribution);
      const maxCare = Math.max(...caretakerDistribution);
      if (maxCare - minCare > 1) {
        isCaretakerEven = false;
      }
    }
    if (isCaretakerEven) {
      evaluation.push(
        `- **世話人配置**: 世話人の配置は比較的均等です (${caretakerDistribution.join(
          ", "
        )})。`
      );
    } else {
      evaluation.push(
        `- **世話人配置**: 世話人の配置は均等ではありません (${caretakerDistribution.join(
          ", "
        )})。`
      );
    }

    // Final output data generation for the current rotation
    tables.forEach((table) => {
      // テーブル内のメンバーをソート
      table.members.sort((a, b) => {
        // 役割ベースの優先順位
        const getRolePriority = (member) => {
          if (member.leader) return 0; // テーブルリーダー (最優先)
          if (member.membership === "ゲスト") return 1; // ゲスト
          // 紹介者の名前がある人がゲストより優先されるように調整 (ゲストグループの固まりを意識)
          // 自身の名前が他のメンバーの紹介者リストに存在するかどうかで判断
          const isIntroducerOfGuestInTable = table.members.some(
            (m) => m.introducer === member.name && m.membership === "ゲスト"
          );
          if (isIntroducerOfGuestInTable) return 2; // ゲストの紹介者

          if (member.caretaker) return 3; // 世話人
          if (member.management) return 4; // 運営部
          return 5; // その他 (最下位)
        };

        const priorityA = getRolePriority(a);
        const priorityB = getRolePriority(b);

        if (priorityA !== priorityB) {
          return priorityA - priorityB;
        }

        // 同じ役割（またはその他）の場合、名前（50音順）でソート
        const collator = new Intl.Collator("ja", { sensitivity: "base" }); // 読み仮名なしで比較
        return collator.compare(a.name, b.name);
      });

      table.members.forEach((member) => {
        let isTableLeaderOutput = member.leader ? 1 : 0;
        let isCaretakerOutput = member.caretaker ? 1 : 0;

        let notes = member.leader ? "テーブルリーダー" : "";

        finalTableAssignment.push({
          rotation: r,
          table_name: `テーブル${table.name}`,
          name: member.name,
          venue: member.venue,
          management: member.management ? 1 : 0,
          leader: isTableLeaderOutput,
          caretaker: isCaretakerOutput,
          introducer: member.introducer || "",
          membership: member.membership,
          notes: notes,
        });
      });
    });
  } // End of rotation loop

  return { tableAssignment: finalTableAssignment, evaluation: evaluation }; // 結果オブジェクトを返す
}

/**
 * Calculates a score for assigning a participant to a given table.
 * Higher score means better fit. This score is used for auto-assignment when multiple tables are valid.
 * @param {Object} participant - 割り当てを検討している参加者
 * @param {Object} tempTable - 参加者を仮割り当てした後のテーブル状態（クローン）
 * @param {Array<Object>} allParticipants - 今月の全ての参加者データ (運営部/世話人の平均計算用)
 * @param {Array<Object>} pastData - 過去のテーブルデータ
 * @param {string} selfVenueName - 自会場名
 * @param {number} totalParticipants - 全参加者数
 * @param {Object} settings - 設定オブジェクト (tableCountなど)
 * @returns {number} 割り当てのスコア
 */
function calculateAssignmentScore(
  participant,
  tempTable,
  allParticipants,
  pastData,
  selfVenueName,
  totalParticipants,
  settings
) {
  let score = 0;
  const targetMinPerTable = Math.floor(totalParticipants / settings.tableCount);
  const targetMaxPerTable =
    targetMinPerTable + (totalParticipants % settings.tableCount > 0 ? 1 : 0);

  // 1. 人数の均等性 (最重要)
  // 目標人数に近いほど高スコア。範囲外は非常に低いスコア。
  const currentTableSize = tempTable.members.length;
  if (
    currentTableSize >= targetMinPerTable &&
    currentTableSize <= targetMaxPerTable
  ) {
    score += 1000; // 目標範囲内は非常に高スコア
    score -=
      Math.min(
        Math.abs(currentTableSize - targetMinPerTable),
        Math.abs(currentTableSize - targetMaxPerTable)
      ) * 10; // 中央に近いほどさらに加点
  } else if (currentTableSize > targetMaxPerTable) {
    score -= 5000; // 人数超過は大きく減点 (isAssignmentValidで弾かれるはずだが、念のため)
  }

  // 2. 運営部・世話人の均等配置 (重要)
  // allParticipants を使用して全体の平均を計算
  const avgManagementPerTable =
    allParticipants.filter((p) => p.management).length / settings.tableCount;
  const avgCaretakerPerTable =
    allParticipants.filter((p) => p.caretaker).length / settings.tableCount;

  score -= Math.abs(tempTable.managementCount - avgManagementPerTable) * 100; // 平均からのずれが大きいほど減点
  score -= Math.abs(tempTable.caretakerCount - avgCaretakerPerTable) * 100;

  // 3. 会場バランス (中程度)
  const selfVenueCount = tempTable.members.filter(
    (m) => m.venue === selfVenueName
  ).length;
  const otherVenueTotalCount = tempTable.members.length - selfVenueCount;
  // 自会場と他会場の比率が均等に近いほど加点
  score -= Math.abs(selfVenueCount - otherVenueTotalCount) * 50;

  // 4. 過去の同席回避 (中程度) - スコアリングで減点方式に戻す
  const hasPastSameSeatConflict = tempTable.autoMembers.some(
    (memberInTable) => {
      // 既存の自動割り当てメンバーとのみ比較
      return pastData.some(
        (past) =>
          past.name === participant.name &&
          pastData.some(
            (past2) =>
              past2.name === memberInTable.name &&
              past.tableNum === past2.tableNum &&
              past.period <= 6
          )
      );
    }
  );
  if (hasPastSameSeatConflict) {
    score -= 100; // 過去同席は減点
  }

  return score;
}

/**
 * メンバーをテーブルに割り当て可能か、絶対的なルールでチェックする関数。
 * isAssignmentValidを通過したテーブルのみがスコアリングの対象となる。
 * @param {Object} participant - 割り当てを検討している参加者
 * @param {Object} table - 割り当て先のテーブル候補
 * @param {Array<Object>} pastData - 過去のテーブルデータ
 * @param {string} selfVenueName - 自会場名
 * @param {number} totalParticipants - 全参加者数
 * @param {Object} settings - 設定オブジェクト (tableCountなど)
 * @param {Set<string>} assignedNamesInRotation - 現在の巡目で既に割り当て済みの名前のセット
 * @returns {boolean} 割り当て可能なら true、そうでなければ false
 */
function isAssignmentValid(
  participant,
  table,
  pastData,
  selfVenueName,
  totalParticipants,
  settings,
  assignedNamesInRotation
) {
  // 1. Rule: 1人の人物は1順につき1回しか登場しないでください (絶対条件)
  if (assignedNamesInRotation.has(participant.name)) {
    return false;
  }

  // 2. Rule: 同じ他会場の方同士は同じテーブルにしない (絶対条件)
  if (
    participant.venue !== selfVenueName &&
    table.otherVenueCounts[participant.venue]
  ) {
    return false;
  }

  // 3. Rule: 人数上限 (絶対条件)
  const targetMaxPerTable =
    Math.floor(totalParticipants / settings.tableCount) +
    (totalParticipants % settings.tableCount > 0 ? 1 : 0);
  if (table.members.length >= targetMaxPerTable) {
    return false;
  }

  // 4. Rule: 各テーブルにテーブルリーダーは1名のみ (絶対条件)
  // 割り当てようとしている参加者がリーダーであり、かつテーブルに既にリーダーがいる場合
  // 割り当てようとしている参加者がリーダーでない場合はこのチェックはスキップされる
  if (participant.leader && table.leaderCount > 0) {
    return false; // テーブルリーダーが複数になるため、割り当て不可
  }

  // 過去の同席回避はスコアリング内で減点されるため、ここでは絶対条件としない。
  // （割り当ては可能だがスコアが低くなる）

  return true;
}

// --- 出力関連関数 ---

/**
 * 生成された席割り結果を指定されたシートに出力する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 出力先シートオブジェクト
 * @param {Array<Object>} assignment - 生成された席割りデータ
 */
function outputTableAssignment(sheet, assignment) {
  if (assignment.length === 0) {
    sheet
      .getRange(OUTPUT_START_CELL)
      .setValue("席割りを生成できませんでした。");
    return;
  }

  // 出力順序の整理: 巡目（昇順）、次にテーブル名（アルファベット順）でソート
  assignment.sort((a, b) => {
    if (a.rotation !== b.rotation) {
      return a.rotation - b.rotation;
    }
    const tableNameA = a.table_name.replace("テーブル", "");
    const tableNameB = b.table_name.replace("テーブル", "");
    return tableNameA.localeCompare(tableNameB); // アルファベット順
  });

  // ヘッダーの定義
  const headers = [
    "巡目",
    "テーブル名",
    "名前",
    "会場",
    "運営部",
    "テーブルリーダー",
    "世話人",
    "紹介者",
    "会員区分",
    "備考",
  ];
  const outputData = [headers];

  // データを行として追加
  assignment.forEach((row) => {
    outputData.push([
      row.rotation,
      row.table_name,
      row.name,
      row.venue,
      row.management,
      row.leader,
      row.caretaker,
      row.introducer,
      row.membership,
      row.notes,
    ]);
  });

  // スプレッドシートへの書き込み
  const startRow = parseInt(OUTPUT_START_CELL.substring(1));
  const startCol = sheet.getRange(OUTPUT_START_CELL).getColumn();
  const range = sheet.getRange(
    startRow,
    startCol,
    outputData.length,
    headers.length
  );
  range.setValues(outputData);
  sheet.autoResizeColumns(startCol, headers.length);
}

/**
 * 評価コメントを指定されたシートに出力する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 出力先シートオブジェクト
 * @param {Array<string>} evaluation - 評価コメントの配列
 */
function outputEvaluation(sheet, evaluation) {
  if (evaluation.length === 0) {
    sheet
      .getRange(EVALUATION_START_CELL)
      .setValue("評価を生成できませんでした。");
    return;
  }
  const startRow = parseInt(EVALUATION_START_CELL.substring(1));
  const startCol = sheet.getRange(EVALUATION_START_CELL).getColumn();

  evaluation.forEach((line, index) => {
    sheet.getRange(startRow + index, startCol).setValue(line);
  });
  sheet.autoResizeColumns(startCol, 1); // 評価コメント列のみ自動調整
}

// --- 過去データ登録関連関数 ---

/**
 * 「今月の席割」シートのデータを「過去の参加者」シートに登録する。
 * 既に今月のデータが存在する場合はエラーメッセージを出して処理終了。
 */
function registerCurrentAssignmentToPastData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  const pastDataSheet = ss.getSheetByName(PAST_TABLE_DATA_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const ui = SpreadsheetApp.getUi(); // Get UI instance for alerts

  // 処理開始メッセージ
  // OK_CANCELを使うことで、処理開始前にユーザーがキャンセルできるようにする
  const confirmDialog = ui.alert(
    "処理中...",
    "過去データへの登録処理を開始します。よろしいですか？",
    ui.ButtonSet.OK_CANCEL
  );
  if (confirmDialog === ui.Button.CANCEL) {
    outputSheet.getRange(CHECKBOX_REGISTER_PAST_DATA_CELL).setValue(false);
    return;
  }

  try {
    Logger.log("--- 過去データ登録処理開始 ---");

    // 必要なシートが存在するかチェック
    if (!outputSheet || !pastDataSheet || !settingsSheet) {
      throw new Error(
        "必要なシートが見つかりません。「今月の席割」「過去の参加者」「設定」を確認してください。"
      );
    }

    const settings = getSettings(settingsSheet);
    const currentMonthNum = settings.currentMonth; // YYYYMM format for the current month

    // Validate '今月' setting
    if (
      !currentMonthNum ||
      typeof currentMonthNum !== "number" ||
      String(currentMonthNum).length !== 6
    ) {
      throw new Error(
        "「設定」シートの「今月」がYYYYMM形式の数値で正しく設定されていません。登録処理を中断します。"
      );
    }

    // --- 「今月の席割」シートからデータを取得 ---
    // データ範囲をA6から最終行・最終列まで正確に指定
    const outputLastRow = outputSheet.getLastRow();
    const outputLastCol = outputSheet.getLastColumn();

    if (outputLastRow < parseInt(OUTPUT_START_CELL.substring(1))) {
      // A6より下にデータがない場合
      throw new Error(
        "「今月の席割」シートに登録すべきデータがありません。A6セル以降に席割りを生成してください。"
      );
    }

    const outputDataRange = outputSheet.getRange(
      parseInt(OUTPUT_START_CELL.substring(1)),
      1,
      outputLastRow - parseInt(OUTPUT_START_CELL.substring(1)) + 1,
      outputLastCol
    );
    const outputValues = outputDataRange.getValues();
    Logger.log("出力シートの読み取り範囲: " + outputDataRange.getA1Notation());

    // ヘッダーが正しく取得できているか確認 (outputValuesの最初の行がヘッダー)
    const outputHeaders = outputValues[0];
    if (outputHeaders.length === 0 || outputHeaders[0] !== "巡目") {
      throw new Error(
        "「今月の席割」シートのヘッダーが不正です。A6セルがヘッダー行であることを確認してください。"
      );
    }

    const outputAssignmentRows = outputValues.slice(1); // データ部分のみ
    Logger.log(
      "出力シートから取得したデータ行数 (ヘッダー除く): " +
        outputAssignmentRows.length
    );

    const outputColIdxRotation = outputHeaders.indexOf("巡目");
    const outputColIdxTableName = outputHeaders.indexOf("テーブル名");
    const outputColIdxName = outputHeaders.indexOf("名前");
    const outputColIdxVenue = outputHeaders.indexOf("会場");

    if (
      outputColIdxRotation === -1 ||
      outputColIdxTableName === -1 ||
      outputColIdxName === -1 ||
      outputColIdxVenue === -1
    ) {
      throw new Error(
        "「今月の席割」シートのヘッダーが不完全です（巡目, テーブル名, 名前, 会場が必須）。登録処理を中断します。"
      );
    }

    // --- 「過去の参加者」シートのヘッダーを確認し、列インデックスを取得 ---
    const pastHeadersRange = pastDataSheet.getRange(
      1,
      1,
      1,
      pastDataSheet.getLastColumn()
    );
    const pastHeaders = pastHeadersRange.getValues()[0];

    const pastColIdxHoldingMonth = pastHeaders.indexOf("開催月");
    const pastColIdxTableNum = pastHeaders.indexOf("卓番");
    const pastColIdxName = pastHeaders.indexOf("名前");
    const pastColIdxVenue = pastHeaders.indexOf("所属会場");

    if (
      pastColIdxHoldingMonth === -1 ||
      pastColIdxTableNum === -1 ||
      pastColIdxName === -1 ||
      pastColIdxVenue === -1
    ) {
      throw new Error(
        "「過去の参加者」シートのヘッダーが不完全です（開催月, 卓番, 名前, 所属会場が必須）。登録処理を中断します。"
      );
    }

    // --- 同月の既存データチェック（削除処理は行わない） ---
    const allPastData = pastDataSheet.getDataRange().getValues();
    if (allPastData.length > 1) {
      // If there's data beyond headers
      for (let i = 1; i < allPastData.length; i++) {
        if (allPastData[i][pastColIdxHoldingMonth] == currentMonthNum) {
          throw new Error(
            `「過去の参加者」シートに今月のデータ（${currentMonthNum}）がすでに存在します。二重登録を避けるため処理を中断します。`
          );
        }
      }
    }

    // --- 新しいデータを整形して登録 ---
    const newPastDataRows = [];
    const numPastCols =
      Math.max(
        pastColIdxHoldingMonth,
        pastColIdxTableNum,
        pastColIdxName,
        pastColIdxVenue
      ) + 1;

    if (outputAssignmentRows.length === 0) {
      // この場合もエラーではなく、完了メッセージを出す
      ui.alert(
        "登録完了",
        `「今月の席割」シートから登録すべき有効なデータが見つかりませんでした。`,
        ui.ButtonSet.OK
      );
      return; // 処理をここで終了
    }

    for (const row of outputAssignmentRows) {
      // 出力シートから必要なデータを取得
      const tableName = String(row[outputColIdxTableName]).replace(
        "テーブル",
        ""
      ); // Convert "テーブルA" to "A"
      const name = String(row[outputColIdxName]).trim();
      const venue = String(row[outputColIdxVenue]).trim();

      let newRow = new Array(numPastCols).fill(""); // Initialize with empty strings

      newRow[pastColIdxHoldingMonth] = currentMonthNum; // Use the current month from settings
      newRow[pastColIdxTableNum] = tableName;
      newRow[pastColIdxName] = name;
      newRow[pastColIdxVenue] = venue;

      newPastDataRows.push(newRow);
    }

    // データを追加する開始行 (既存データの最終行の次)
    const appendRowStart = pastDataSheet.getLastRow() + 1;
    pastDataSheet
      .getRange(
        appendRowStart,
        1,
        newPastDataRows.length,
        newPastDataRows[0].length
      )
      .setValues(newPastDataRows);

    // 転記が確実に反映されるのを待つ
    SpreadsheetApp.flush();

    // 転記が確実に完了した後にダイアログを表示
    ui.alert(
      "登録完了",
      `今月の席割りデータ (${currentMonthNum}) を過去の参加者シートに登録しました。`,
      ui.ButtonSet.OK
    );
    Logger.log("--- 過去データ登録処理正常終了 ---");
  } catch (e) {
    // エラー発生時のログ出力とダイアログ表示
    Logger.log("--- 過去データ登録処理エラー終了 ---");
    Logger.log("エラー: " + e.message + " (スタック: " + e.stack + ")");
    ui.alert(
      "登録エラー",
      "エラーが発生しました: " + e.message + "\n処理を中断します。",
      ui.ButtonSet.OK
    );
  } finally {
    // 処理の成功/失敗に関わらず、チェックボックスを自動でOFFに戻す
    outputSheet.getRange(CHECKBOX_REGISTER_PAST_DATA_CELL).setValue(false);
  }
}
