// バージョン: V6.40 (症例評価A-F・初回バイタル・初回血液検査の未使用列をタイムラインJSONから自動集計)
// ※このファイルはリポジトリ管理用のミラーです。実際の反映には
//   script.google.com のプロジェクトに貼り付けて「新しいデプロイ」または
//   既存デプロイの「新バージョン」として公開する必要があります。

function sendAdminEmail(subject, body) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName('マスタ＿管理') || ss.getSheetByName('マスタ_管理');
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length > 1 && data[1].length > 0) {
      var email = data[1][0];
      if (email && String(email).indexOf('@') !== -1) {
        MailApp.sendEmail({
          to: String(email).trim(),
          subject: "[小児救急アプリ V6.23] " + subject,
          body: "システムからの通知:\n\n" + body
        });
      }
    }
  } catch (e) {
    Logger.log("管理者メール送信失敗: " + e.toString());
  }
}

function safeGet(row, hm, key) {
  if (hm && hm.hasOwnProperty(key) && hm[key] < row.length) {
    var val = row[hm[key]];
    return (val === null || val === undefined) ? '' : val;
  }
  return '';
}

// スプレッドシートが日付入力を自動でDate型に変換した場合、表示用にYYYY/MM/DD文字列へ正規化する
function formatDateVal(val) {
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd");
  return val;
}

// <input type="date">にそのままセットできるYYYY-MM-DD形式に正規化する
// （セルがDate型化されているとgetValuesはISO日時文字列を返し、date inputへの
//   代入が無効値として無視され、次回保存時に空欄で上書きされる原因になっていた）
function formatDateForInput(val) {
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Tokyo", "yyyy-MM-dd");
  if (!val) return '';
  return String(val).split('T')[0];
}

// <input type="time">にそのままセットできるHH:mm形式に正規化する
// （開始時間・終了時間のセルがTime型化されると同様にDateで返ってくるため）
function formatTimeForInput(val) {
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Tokyo", "HH:mm");
  if (!val) return '';
  var m = String(val).match(/(\d{1,2}):(\d{2})/);
  return m ? (m[1].length === 1 ? '0' + m[1] : m[1]) + ':' + m[2] : String(val);
}

function doGet(e) {
  var callback = e.parameter ? e.parameter.callback : null;
  try {
    var data = getMasterData();
    var result = JSON.stringify(data);
    if (callback) return ContentService.createTextOutput(callback + '(' + result + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    var errObj = JSON.stringify({ status: "error", message: err.toString() });
    sendAdminEmail("doGet内エラー", "エラー詳細:\n" + err.toString() + "\n\nスタックトレース:\n" + err.stack);
    if (callback) return ContentService.createTextOutput(callback + '(' + errObj + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(errObj).setMimeType(ContentService.MimeType.JSON);
  }
}

function getHeaderMap(sheet) {
  var map = {};
  if (!sheet) return map;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return map;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) map[String(headers[i]).trim()] = i;
  }
  return map;
}

// 振り返りのタイムラインJSONから、スプレッドシート上で一括処理・
// ソート・フィルタしやすいよう「初回」の値だけを抜き出す。
// ・症例評価A〜F: 最も時刻の早いABCDEF評価(type==='eval')の各項目
// ・初回バイタル各項目: 時刻順に見て、その項目が最初に入力された値
//   （1件のバイタル登録に全項目が揃っているとは限らないため、
//   　項目ごとに独立して「最初に値が入った時刻」を探す）
// ・初回血液検査各項目: 同様に項目（検査名）ごとに最初の値を探す
function extractDebriefSummaryFields(timeline) {
  var result = {
    evalA: '', evalB: '', evalC: '', evalD: '', evalE: '', evalF: '',
    hr: '', spo2: '', rr: '', bpSys: '', bpDia: '', bt: '', loc: '', pews: '',
    ph: '', pco2: '', po2: '', hco3: '', be: '', lac: '', wbc: '', hb: '', plt: '', crp: ''
  };
  if (!timeline || timeline.length === 0) return result;

  var evalItems = timeline.filter(function(t) { return t.type === 'eval' && t.time; }).sort(function(a, b) { return String(a.time).localeCompare(String(b.time)); });
  if (evalItems.length > 0) {
    var first = evalItems[0];
    result.evalA = first.a || ''; result.evalB = first.b || ''; result.evalC = first.c || '';
    result.evalD = first.d || ''; result.evalE = first.e || ''; result.evalF = first.f || '';
  }

  var vitals = timeline.filter(function(t) { return t.cat === 'バイタル' && t.time; }).sort(function(a, b) { return String(a.time).localeCompare(String(b.time)); });
  var vitalMap = { hr: 'v-hr', spo2: 'v-spo2', rr: 'v-rr', bpSys: 'v-bps', bpDia: 'v-bpd', bt: 'v-bt', loc: 'v-jcs', pews: 'v-pews' };
  Object.keys(vitalMap).forEach(function(key) {
    var rawKey = vitalMap[key];
    for (var i = 0; i < vitals.length; i++) {
      var v = vitals[i].rawValues && vitals[i].rawValues[rawKey];
      if (v !== undefined && v !== null && v !== '') { result[key] = v; break; }
    }
  });

  var bloodTests = timeline.filter(function(t) { return t.cat === '血液検査' && t.time; }).sort(function(a, b) { return String(a.time).localeCompare(String(b.time)); });
  // マスタ＿検査の正式名称と完全一致させる（"BE (Base Excess)"と
  // "BE(ecf) (細胞外液基剰塩基)"のように前方一致だと紛れる項目があるため）
  var testMap = {
    ph: 'pH', pco2: 'pCO2', po2: 'pO2', hco3: 'HCO3-', be: 'BE (Base Excess)', lac: 'Lactate (乳酸値)',
    wbc: 'WBC (白血球数)', hb: 'Hb (ヘモグロビン)', plt: 'Plt (血小板数)', crp: 'CRP'
  };
  Object.keys(testMap).forEach(function(key) {
    var rawKey = 'test_' + testMap[key];
    for (var i = 0; i < bloodTests.length; i++) {
      var raw = bloodTests[i].rawValues || {};
      if (raw[rawKey] !== undefined && raw[rawKey] !== null && raw[rawKey] !== '') { result[key] = raw[rawKey]; break; }
    }
  });

  return result;
}

// シートに指定した列見出しが無ければ、末尾に新しい列を追加してから
// 最新のヘッダーマップを返す（既存シートへの後付け列追加用）
function ensureColumn(sheet, columnName) {
  var hm = getHeaderMap(sheet);
  if (hm.hasOwnProperty(columnName)) return hm;
  var lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1).setValue(columnName);
  return getHeaderMap(sheet);
}

function createRowData(headerMap, dataDict) {
  var row = new Array(Object.keys(headerMap).length).fill("");
  for (var key in dataDict) {
    if (headerMap.hasOwnProperty(key)) {
      row[headerMap[key]] = dataDict[key];
    }
  }
  return row;
}

// シートが無ければヘッダー付きで自動作成する（ポータル用シート向け）
function ensureSheetWithHeaders(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function doPost(e) {
  try {
    var parsedPayload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. エラー報告
    if (parsedPayload.action === "report_client_error") {
      sendAdminEmail("フロントエンドエラー報告 (" + (parsedPayload.source || "不明") + ")", parsedPayload.errorInfo || "詳細不明");
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 動的マスタ追加
    if (parsedPayload.action === "add_to_master") {
      var target = parsedPayload.target; var val1 = parsedPayload.val1; var val2 = parsedPayload.val2 || "";
      if (!val1) throw new Error("登録する値が空です");
      var sheetAdmin = ss.getSheetByName('マスタ＿管理') || ss.getSheetByName('マスタ_管理');

      if (target === "keyword") {
        if (!sheetAdmin) throw new Error("マスタ管理シートが見つかりません");
        sheetAdmin.getRange(Math.max(sheetAdmin.getLastRow(), 1) + 1, 2).setValue(val1);
      }
      else if (target === "hospital") {
        if (!sheetAdmin) throw new Error("マスタ管理シートが見つかりません");
        var lastRow = Math.max(sheetAdmin.getLastRow(), 1);
        sheetAdmin.getRange(lastRow + 1, 3).setValue(val2 || "その他"); sheetAdmin.getRange(lastRow + 1, 4).setValue(val1);
      }
      else if (target === "ems") {
        if (!sheetAdmin) throw new Error("マスタ管理シートが見つかりません");
        var lastRow = Math.max(sheetAdmin.getLastRow(), 1);
        sheetAdmin.getRange(lastRow + 1, 5).setValue(val2 || "その他地域"); sheetAdmin.getRange(lastRow + 1, 6).setValue(val1);
      }
      else if (target === "staff") {
        var sheetStaff = ss.getSheetByName('マスタ＿スタッフ') || ss.getSheetByName('マスタースタッフ') || ss.getSheetByName('マスタ_スタッフ');
        if (!sheetStaff) throw new Error("スタッフマスタが見つかりません");
        var hm = getHeaderMap(sheetStaff);
        var newRow = createRowData(hm, { 'フルネーム(選択用)': val1, '表示名': val1, '氏名': val1, '姓': val1, '職種': val2 });
        sheetStaff.appendRow(newRow);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. 事案登録 (作戦盤用)
    if (parsedPayload.action === "peds_submit") {
      var dbSheet = ss.getSheetByName('事案アーカイブ');
      if (!dbSheet) throw new Error("事案アーカイブシートが見つかりません");
      var hm = ensureColumn(dbSheet, '完全復元用JSON');
      var timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
      var p = parsedPayload.payload || {}; var pd = p.data || {}; var roles = p.staffRoles || {};

      var pedsData = dbSheet.getDataRange().getValues(); var pedsTargetRow = -1;
      if (pd["要請番号"]) {
        for (var pr = 1; pr < pedsData.length; pr++) { if (safeGet(pedsData[pr], hm, '要請番号') === pd["要請番号"]) { pedsTargetRow = pr + 1; break; } }
      }
      var existingPedsRow = (pedsTargetRow > -1) ? pedsData[pedsTargetRow - 1] : null;

      // submit_debriefと同様、送信値が空の項目は既存の保存済み値を維持する。
      // 復元処理が全項目をカバーしきれていない場合でも、再送信のたびに
      // 既存データが空欄で上書き消去されるのを防ぐための安全策。
      function keepIfEmptyPeds(newVal, colName) {
        if (newVal !== '' && newVal !== undefined && newVal !== null) return newVal;
        if (existingPedsRow) { var old = safeGet(existingPedsRow, hm, colName); if (old) return old; }
        return '';
      }

      // 復元漏れがあっても次回開いたときに完全復元できるよう、送信された
      // appStateの生データをJSON列にもフル保存しておく（画面の各表示文字列は
      // 人間が読む要約であり、そこから逆算復元するのは不正確・不安定なため）
      var fullStateJson = JSON.stringify(p);

      var dataDict = {
        '分野区分': '小児', '保存日時': timestamp, '要請番号': pd["要請番号"] || '',
        '搬入日': keepIfEmptyPeds(pd["日付"] || '', '搬入日'),
        '事案種別': keepIfEmptyPeds(pd["事案種別"] || '', '事案種別'),
        '紹介元医療機関': keepIfEmptyPeds(pd["紹介元"] || '', '紹介元医療機関'),
        '年齢/月齢': keepIfEmptyPeds(p.ageText || '', '年齢/月齢'),
        '性別': keepIfEmptyPeds(pd["性別"] || '', '性別'),
        '計算体重': keepIfEmptyPeds(p.weight || '', '計算体重'),
        '目安身長': keepIfEmptyPeds(p.height || '', '目安身長'),
        'キーワード': keepIfEmptyPeds((p.keywords || []).join(', '), 'キーワード'),
        '概要・経過': keepIfEmptyPeds(pd["概要"] || '', '概要・経過'),
        '搬入前処置': keepIfEmptyPeds(pd["搬入前処置"] || '', '搬入前処置'),
        'AMPL': keepIfEmptyPeds(pd["AMPL"] || '', 'AMPL'),
        'PAT': keepIfEmptyPeds(pd["PAT"] || '', 'PAT'),
        'バイタル': keepIfEmptyPeds(pd["バイタル"] || '', 'バイタル'),
        'PEWS': keepIfEmptyPeds(pd["PEWS"] || '', 'PEWS'),
        '統括': keepIfEmptyPeds((roles.leader || []).join(', '), '統括'),
        '気道管理': keepIfEmptyPeds((roles.airway || []).join(', '), '気道管理'),
        '胸骨圧迫': keepIfEmptyPeds((roles.cpr || []).join(', '), '胸骨圧迫'),
        'ルート・薬剤': keepIfEmptyPeds((roles.route || []).join(', '), 'ルート・薬剤'),
        '記録': keepIfEmptyPeds((roles.record || []).join(', '), '記録'),
        'その他役割': keepIfEmptyPeds((roles.other || []).join(', '), 'その他役割'),
        '想定シナリオ・薬剤物品詳細': keepIfEmptyPeds(pd["プロトコル詳細"] || '', '想定シナリオ・薬剤物品詳細'),
        '完全復元用JSON': fullStateJson
      };
      // 要請番号が既存行と一致する場合は追記ではなく上書き更新する。
      // 常にappendRowしていたため、同じ事案を開き直して再送信するだけで
      // （URLからの復元＋再送信、二重送信など）事案アーカイブ検索に
      // 全く同じ内容の行が2件以上並ぶ原因になっていた。
      var newRow = createRowData(hm, dataDict);
      if (pedsTargetRow > -1) dbSheet.getRange(pedsTargetRow, 1, 1, newRow.length).setValues([newRow]);
      else dbSheet.appendRow(newRow);
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. 事案アーカイブ検索
    if (parsedPayload.action === "fetch_history") {
      var dbSheet = ss.getSheetByName('事案アーカイブ');
      if (!dbSheet) throw new Error("事案アーカイブが見つかりません");
      var hm = getHeaderMap(dbSheet); var data = dbSheet.getDataRange().getValues(); var results = [];
      var sMonth = parsedPayload.month || ""; var sStaff = parsedPayload.staff || ""; var sKw = parsedPayload.keyword || "";
      var sStatus = parsedPayload.debriefStatus || ""; // "完了" | "未完了" | ""(全て)

      // 振り返り状況を要請番号ごとに1回だけ集計（完了ステータスの事案IDセットを作成）
      // ついでに、振り返りのタイムラインJSONから実際に投与された薬剤名も
      // 事案ごとに集計しておく（検索コンソールで「使用した薬剤で検索」を
      // 可能にするため。用量までは検索対象にせず薬剤名のみで十分という判断）。
      var debriefStatusMap = {};
      var debriefDrugsMap = {};
      try {
        var debriefSheetForStatus = ss.getSheetByName('振り返りアーカイブ');
        if (debriefSheetForStatus) {
          var hmDebriefStatus = getHeaderMap(debriefSheetForStatus);
          var debriefValsForStatus = debriefSheetForStatus.getDataRange().getValues();
          for (var j = 1; j < debriefValsForStatus.length; j++) {
            var dRowS = debriefValsForStatus[j];
            var dIdS = safeGet(dRowS, hmDebriefStatus, '要請番号');
            var dStatusS = String(safeGet(dRowS, hmDebriefStatus, 'ステータス')).trim();
            if (dIdS) debriefStatusMap[dIdS] = dStatusS === '完了' ? '完了' : '未完了';
            if (dIdS) {
              try {
                var tlJson = safeGet(dRowS, hmDebriefStatus, 'タイムラインJSON');
                if (tlJson) {
                  var tlParsed = JSON.parse(tlJson);
                  var drugNames = tlParsed.filter(function(t) { return t.cat === '薬剤' && t.contentText; })
                    .map(function(t) { return String(t.contentText).split(' ')[0]; });
                  if (drugNames.length > 0) {
                    var existing = debriefDrugsMap[dIdS] ? debriefDrugsMap[dIdS].split(',') : [];
                    var merged = existing.concat(drugNames).filter(function(v, idx, arr) { return arr.indexOf(v) === idx; });
                    debriefDrugsMap[dIdS] = merged.join(',');
                  }
                }
              } catch (eTl) { /* 個別事案のJSON壊れは無視して続行 */ }
            }
          }
        }
      } catch (eStatus) { Logger.log("振り返り状況集計エラー: " + eStatus.toString()); }

      for (var i = 1; i < data.length; i++) {
        var row = data[i]; var id = safeGet(row, hm, '要請番号'); var rawDate = safeGet(row, hm, '搬入日'); var date = formatDateVal(rawDate); var summary = safeGet(row, hm, '概要・経過'); var keywords = safeGet(row, hm, 'キーワード');
        if (sKw && id.indexOf(sKw) === -1 && summary.indexOf(sKw) === -1 && keywords.indexOf(sKw) === -1) continue;
        // 「搬入日」セルがDate型化されている場合、rawDate(Dateオブジェクト)を
        // そのままString()化すると toString() 形式("Wed Aug 02 2026...")になり、
        // "yyyy-MM"形式の検索キーワードと絶対に一致しないため月検索が機能していなかった。
        // 既にyyyy/MM/dd文字列へ正規化済みの date を使う。
        if (sMonth && String(date).replace(/\//g, '-').indexOf(sMonth) === -1) continue;
        // 「対応スタッフ」検索：sStaffパラメータはこれまで受け取るだけで
        // 一切フィルタに使われておらず、常に無効化されていた。
        // ラベル・オートコンプリートの内容から、ブリーフィングで役割分担された
        // スタッフ（統括/気道管理/胸骨圧迫/ルート・薬剤/記録）を検索する項目と判断し実装。
        if (sStaff) {
          var rolesStr = [safeGet(row, hm, '統括'), safeGet(row, hm, '気道管理'), safeGet(row, hm, '胸骨圧迫'), safeGet(row, hm, 'ルート・薬剤'), safeGet(row, hm, '記録')].join(',');
          if (rolesStr.indexOf(sStaff) === -1) continue;
        }
        var debriefStatus = debriefStatusMap.hasOwnProperty(id) ? debriefStatusMap[id] : '未入力';
        if (sStatus && sStatus !== debriefStatus) continue;
        results.push({
          id: id, date: date, age: safeGet(row, hm, '年齢/月齢'), weight: safeGet(row, hm, '計算体重'), sex: safeGet(row, hm, '性別'), summary: summary, keywords: keywords, preTx: safeGet(row, hm, '搬入前処置'), protocol: safeGet(row, hm, '想定シナリオ・薬剤物品詳細'), vitals: safeGet(row, hm, 'バイタル'), pat: safeGet(row, hm, 'PAT'), caseType: safeGet(row, hm, '事案種別'), facility: safeGet(row, hm, '紹介元医療機関'), debriefStatus: debriefStatus,
          roles: { leader: safeGet(row, hm, '統括'), airway: safeGet(row, hm, '気道管理'), cpr: safeGet(row, hm, '胸骨圧迫'), route: safeGet(row, hm, 'ルート・薬剤'), record: safeGet(row, hm, '記録') },
          fullState: safeGet(row, hm, '完全復元用JSON'),
          usedDrugs: debriefDrugsMap[id] || ''
        });
      }
      results.reverse();
      if (results.length === 0) return ContentService.createTextOutput(JSON.stringify({ status: "not_found" })).setMimeType(ContentService.MimeType.JSON);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", results: results })).setMimeType(ContentService.MimeType.JSON);
    }

    // 5. デブリーフィング取得
    if (parsedPayload.action === "fetch_debrief") {
      var dbSheet = ss.getSheetByName('振り返りアーカイブ');
      if (!dbSheet) throw new Error("振り返りアーカイブが見つかりません");
      var hm = getHeaderMap(dbSheet); var data = dbSheet.getDataRange().getValues();
      var debriefData = {};
      for (var i = data.length - 1; i >= 1; i--) {
        if (safeGet(data[i], hm, '要請番号') === parsedPayload.id) {
          debriefData = { actualDiff: safeGet(data[i], hm, '想定との相違点'), gapBad: safeGet(data[i], hm, 'ギャップ課題'), gapGood: safeGet(data[i], hm, 'ギャップ良かった点'), gapMaster: safeGet(data[i], hm, 'マスタ改修提案'), team: safeGet(data[i], hm, 'チーム連携評価'), action: safeGet(data[i], hm, '次回アクションプラン'), dDate: formatDateForInput(safeGet(data[i], hm, '開催日')), dTime: formatTimeForInput(safeGet(data[i], hm, '開始時間')), dEndTime: formatTimeForInput(safeGet(data[i], hm, '終了時間')), dDuration: safeGet(data[i], hm, '所要時間'), dPlace: safeGet(data[i], hm, '場所'), timeline: safeGet(data[i], hm, 'タイムラインJSON'), status: safeGet(data[i], hm, 'ステータス'), summary: safeGet(data[i], hm, '事案概要'), patientInfo: safeGet(data[i], hm, '患者情報'), recorder: safeGet(data[i], hm, '記録者'), actualStaffs: safeGet(data[i], hm, '実際の対応スタッフ'), debriefStaffs: safeGet(data[i], hm, '参加スタッフ') }; break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", debriefData: debriefData })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6. デブリーフィング保存
    if (parsedPayload.action === "submit_debrief") {
      var dbSheet = ss.getSheetByName('振り返りアーカイブ');
      if (!dbSheet) throw new Error("振り返りアーカイブが見つかりません");
      var hm = getHeaderMap(dbSheet); var p = parsedPayload.payload || {};
      var data = dbSheet.getDataRange().getValues(); var targetRow = -1;
      for (var i = 1; i < data.length; i++) { if (safeGet(data[i], hm, '要請番号') === p.id) { targetRow = i + 1; break; } }
      var existingRow = (targetRow > -1) ? data[targetRow - 1] : null;

      // 画面側で復元漏れ等により項目が空のまま送信された場合でも、
      // 既存の保存済み内容を空欄で上書きして消してしまわないよう、
      // 送信値が空の項目はスプレッドシート上の既存値を維持する。
      // （タイムスタンプ・要請番号・ステータスは常に今回の送信値で確定させる）
      function keepIfEmpty(newVal, colName) {
        if (newVal !== '' && newVal !== undefined && newVal !== null) return newVal;
        if (existingRow) { var old = safeGet(existingRow, hm, colName); if (old) return old; }
        return '';
      }
      var timelineJson = (p.timeline && p.timeline.length > 0)
        ? JSON.stringify(p.timeline)
        : keepIfEmpty('', 'タイムラインJSON');

      // 症例評価A〜F・初回バイタル・初回血液検査の各列は、以前から
      // シートに用意されていたが、これまでコードから一切書き込まれて
      // いなかった（未使用の空列）。タイムラインJSONの中に同じ情報が
      // 既にあるため、スプレッドシート単体でも読みやすく・一括処理
      // しやすいよう、保存のたびにここから要約して書き込む。
      var summaryFields = extractDebriefSummaryFields(p.timeline || []);

      var dataDict = {
        'タイムスタンプ': Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"),
        '要請番号': p.id || '',
        '発生日': keepIfEmpty(p.date || '', '発生日'),
        '患者情報': keepIfEmpty(p.patientInfo || '', '患者情報'),
        '参加スタッフ': keepIfEmpty(p.debriefStaffs || '', '参加スタッフ'),
        '事案概要': keepIfEmpty(p.summary || '', '事案概要'),
        'タイムラインJSON': timelineJson,
        '記録者': keepIfEmpty(p.recorder || '', '記録者'),
        '実際の対応スタッフ': keepIfEmpty(p.actualStaffs || '', '実際の対応スタッフ'),
        '想定との相違点': keepIfEmpty(p.actualDiff || '', '想定との相違点'),
        'ギャップ課題': keepIfEmpty(p.gapBad || '', 'ギャップ課題'),
        'ギャップ良かった点': keepIfEmpty(p.gapGood || '', 'ギャップ良かった点'),
        'チーム連携評価': keepIfEmpty(p.team || '', 'チーム連携評価'),
        'マスタ改修提案': keepIfEmpty(p.gapMaster || '', 'マスタ改修提案'),
        '次回アクションプラン': keepIfEmpty(p.action || '', '次回アクションプラン'),
        '開催日': keepIfEmpty(p.dDate || '', '開催日'),
        '開始時間': keepIfEmpty(p.dTime || '', '開始時間'),
        '終了時間': keepIfEmpty(p.dEndTime || '', '終了時間'),
        '所要時間': keepIfEmpty(p.dDuration || '', '所要時間'),
        '場所': keepIfEmpty(p.dPlace || '', '場所'),
        'ステータス': p.status || '',
        '症例評価A': keepIfEmpty(summaryFields.evalA, '症例評価A'),
        '症例評価B': keepIfEmpty(summaryFields.evalB, '症例評価B'),
        '症例評価C': keepIfEmpty(summaryFields.evalC, '症例評価C'),
        '症例評価D': keepIfEmpty(summaryFields.evalD, '症例評価D'),
        '症例評価E': keepIfEmpty(summaryFields.evalE, '症例評価E'),
        '症例評価F': keepIfEmpty(summaryFields.evalF, '症例評価F'),
        '初回_HR': keepIfEmpty(summaryFields.hr, '初回_HR'),
        '初回_SpO2': keepIfEmpty(summaryFields.spo2, '初回_SpO2'),
        '初回_RR': keepIfEmpty(summaryFields.rr, '初回_RR'),
        '初回_BP収縮': keepIfEmpty(summaryFields.bpSys, '初回_BP収縮'),
        '初回_BP拡張': keepIfEmpty(summaryFields.bpDia, '初回_BP拡張'),
        '初回_BT': keepIfEmpty(summaryFields.bt, '初回_BT'),
        '初回_意識レベル': keepIfEmpty(summaryFields.loc, '初回_意識レベル'),
        '初回_PEWS': keepIfEmpty(summaryFields.pews, '初回_PEWS'),
        '初回_pH': keepIfEmpty(summaryFields.ph, '初回_pH'),
        '初回_pCO2': keepIfEmpty(summaryFields.pco2, '初回_pCO2'),
        '初回_pO2': keepIfEmpty(summaryFields.po2, '初回_pO2'),
        '初回_HCO3': keepIfEmpty(summaryFields.hco3, '初回_HCO3'),
        '初回_BE': keepIfEmpty(summaryFields.be, '初回_BE'),
        '初回_Lac': keepIfEmpty(summaryFields.lac, '初回_Lac'),
        '初回_WBC': keepIfEmpty(summaryFields.wbc, '初回_WBC'),
        '初回_Hb': keepIfEmpty(summaryFields.hb, '初回_Hb'),
        '初回_Plt': keepIfEmpty(summaryFields.plt, '初回_Plt'),
        '初回_CRP': keepIfEmpty(summaryFields.crp, '初回_CRP')
      };
      var newRow = createRowData(hm, dataDict);
      if (targetRow > -1) dbSheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
      else dbSheet.appendRow(newRow);
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 7. AIカルテ解析
    if (parsedPayload.action === "analyze_karte") {
      var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!apiKey) throw new Error("APIキーが設定されていません。");
      var target = parsedPayload.target;
      // 振り返り内容のAI下書き要約：他のtargetと異なり画像OCRではなく、
      // 入力済みタイムライン・医学的評価から振り返り文章の下書きを生成する
      if (target === "debrief_summary") {
        var summaryPrompt = "あなたは小児救急のインシデントレビュー（デブリーフィング）を支援するアシスタントです。"
          + "以下のタイムライン記録と医学的評価（ABCDEFアプローチ）をもとに、振り返りシートの「詳細分析・チーム連携評価」欄にそのまま使える"
          + "日本語の下書き文章を200〜350字程度で作成してください。事実に基づいて淡々と記述し、断定的な医学的判断や責任追及の表現は避け、"
          + "良かった点と気付いた点（改善余地）をバランス良く含めてください。マークダウン不可。出力は次のJSON形式のみ: {\"summary\": \"...\"}\n\n"
          + "【タイムライン】\n" + JSON.stringify(parsedPayload.timeline || []) + "\n\n"
          + "【医学的評価(ABCDEF)】\n" + JSON.stringify(parsedPayload.evalBlocks || []);
        // Gemini側が「高負荷」で一時的に失敗することがあり(This model is
        // currently experiencing high demand...)、その場合は少し待って
        // 再試行すれば通ることが多いため、最大2回まで自動リトライする。
        var summaryJson = null; var summaryLastErr = null;
        for (var attempt = 1; attempt <= 2; attempt++) {
          var summaryRes = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=" + apiKey, {
            method: "post", contentType: "application/json", muteHttpExceptions: true,
            payload: JSON.stringify({ contents: [{ parts: [{ text: summaryPrompt }] }], generationConfig: { temperature: 0.3, responseMimeType: "application/json" } })
          });
          summaryJson = JSON.parse(summaryRes.getContentText());
          if (!summaryJson.error) { summaryLastErr = null; break; }
          summaryLastErr = summaryJson.error;
          if (attempt < 2 && /high demand|overloaded|503/i.test(summaryLastErr.message || '')) { Utilities.sleep(2000); continue; }
          break;
        }
        if (summaryLastErr) throw new Error(summaryLastErr.message);
        var summaryClean = summaryJson.candidates[0].content.parts[0].text;
        var summaryMatch = summaryClean.match(/\{[\s\S]*\}/);
        if (summaryMatch) summaryClean = summaryMatch[0];
        return ContentService.createTextOutput(JSON.stringify({ status: "success", target: target, data: JSON.parse(summaryClean) })).setMimeType(ContentService.MimeType.JSON);
      }

      var promptText = "あなたは医療記録アシスタントです。提供されたデータから指定JSONのみ出力。マークダウン不可。読み取れない項目はキー自体を省略してください。数値以外の文字（単位・記号）は含めないでください。\n";
      if (target === "timeline") {
        promptText += "【抽出】時間(time)と内容(contentText)の配列。catは'処置','薬剤','記事'。例:[{\"time\":\"14:23\",\"cat\":\"記事\",\"contentText\":\"挿管\"}]";
      } else if (target === "vital") {
        promptText += "【抽出】バイタルサインを次のキーちょうどで出力してください: v_hr(心拍数/HR), v_spo2(SpO2%), v_rr(呼吸数/RR), v_bps(収縮期血圧), v_bpd(拡張期血圧), v_bt(体温)。例:{\"v_hr\":120,\"v_spo2\":98}";
      } else if (target === "blood") {
        var knownNames = getBloodTestNames(ss);
        promptText += "【抽出】この検体検査結果の画像・テキストから、数値が読み取れた項目だけを"
          + "次の候補リストの表記に厳密に一致させて出力してください（候補にない項目は無視）: "
          + knownNames.join('、') + "。\n出力形式は {\"項目名\": 数値, ...} 。項目名は必ず候補リストの文字列をそのまま使うこと。";
      } else {
        promptText += "【抽出】" + target + "に関する数値をJSON形式で出力。";
      }

      var parts = [{ text: promptText }];
      if (parsedPayload.imageText) parts.push({ text: "【テキスト】\n" + parsedPayload.imageText });
      if (parsedPayload.imageBase64) parts.push({ inlineData: { mimeType: parsedPayload.mimeType || "image/jpeg", data: parsedPayload.imageBase64.replace(/^data:image\/[a-z]+;base64,/, "") } });

      var res = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=" + apiKey, {
        method: "post", contentType: "application/json", muteHttpExceptions: true, payload: JSON.stringify({ contents: [{ parts: parts }], generationConfig: { temperature: 0.0, responseMimeType: "application/json" } })
      });
      var json = JSON.parse(res.getContentText()); if (json.error) throw new Error(json.error.message);
      var cleanText = json.candidates[0].content.parts[0].text;
      var jsonMatch = cleanText.match(/\{[\s\S]*\}|\[[\s\S]*\]/);
      if (jsonMatch) cleanText = jsonMatch[0];
      return ContentService.createTextOutput(JSON.stringify({ status: "success", target: target, data: JSON.parse(cleanText) })).setMimeType(ContentService.MimeType.JSON);
    }

    // 8. ポータル情報取得（お知らせ・マニュアル・未完了アラート）
    if (parsedPayload.action === "fetch_portal_data") {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getPortalData(ss) })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    sendAdminEmail("doPost内処理エラー", err.toString() + "\n" + err.stack);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// AI画像解析（血液検査OCR）が結果を照合できるよう、マスタ＿検査シートから
// 血液カテゴリの項目名一覧だけを軽量に取得する
function getBloodTestNames(ss) {
  var names = [];
  var sheetTest = ss.getSheetByName('マスタ＿検査');
  if (!sheetTest) return names;
  var tData = sheetTest.getDataRange().getValues();
  var hMapTest = {}; var dataStartRow = 1;
  for (var i = 0; i < tData.length; i++) {
    if (tData[i].indexOf('大項目') !== -1 || tData[i].indexOf('検査項目名') !== -1) {
      for (var c = 0; c < tData[i].length; c++) { if (tData[i][c]) hMapTest[String(tData[i][c]).trim()] = c; }
      dataStartRow = i + 1; break;
    }
  }
  for (var i = dataStartRow; i < tData.length; i++) {
    var r = tData[i];
    var name = safeGet(r, hMapTest, '検査項目名');
    var category = safeGet(r, hMapTest, '大項目');
    // デブリーフィング側の検査値入力パネルは血液以外（尿・髄液・培養等）も
    // 対象にしたため、OCR照合候補も同じ範囲に合わせる
    if (name && (category === '血液' || category === '尿' || category === '髄液' || category === 'その他')) names.push(String(name).trim());
  }
  return names;
}

// ----------------------------------------------------
// 【ポータル用データ取得】お知らせ・マニュアル・未完了アラート
// ----------------------------------------------------
function getPortalData(ss) {
  var result = { news: [], manual: [], alerts: [] };

  // お知らせ: 日付, 区分, タイトル, 本文, ステータス（「公開」のみ返却・新しい順に最大10件）
  var newsSheet = ensureSheetWithHeaders(ss, 'お知らせ', ['日付', '区分', 'タイトル', '本文', 'ステータス']);
  var hmNews = getHeaderMap(newsSheet);
  var newsVals = newsSheet.getDataRange().getValues();
  for (var i = 1; i < newsVals.length; i++) {
    var r = newsVals[i];
    var title = safeGet(r, hmNews, 'タイトル');
    if (!title) continue;
    var status = String(safeGet(r, hmNews, 'ステータス')).trim();
    if (status && status !== '公開') continue;
    result.news.push({
      date: formatDateVal(safeGet(r, hmNews, '日付')),
      category: safeGet(r, hmNews, '区分'),
      title: title,
      content: safeGet(r, hmNews, '本文')
    });
  }
  result.news.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  result.news = result.news.slice(0, 10);

  // マニュアル: 区分, タイトル, 本文
  var manualSheet = ensureSheetWithHeaders(ss, 'マニュアル', ['区分', 'タイトル', '本文']);
  var hmManual = getHeaderMap(manualSheet);
  var manualVals = manualSheet.getDataRange().getValues();
  for (var i = 1; i < manualVals.length; i++) {
    var r = manualVals[i];
    var mTitle = safeGet(r, hmManual, 'タイトル');
    if (!mTitle) continue;
    result.manual.push({
      category: safeGet(r, hmManual, '区分'),
      title: mTitle,
      content: safeGet(r, hmManual, '本文')
    });
  }

  // 未完了アラート: 事案アーカイブ にあって 振り返りアーカイブ が「完了」でない事案（直近60日・最大5件）
  try {
    var caseSheet = ss.getSheetByName('事案アーカイブ');
    var debriefSheet = ss.getSheetByName('振り返りアーカイブ');
    if (caseSheet && debriefSheet) {
      var hmCase = getHeaderMap(caseSheet);
      var hmDebrief = getHeaderMap(debriefSheet);
      var debriefVals = debriefSheet.getDataRange().getValues();
      var completedIds = {};
      for (var i = 1; i < debriefVals.length; i++) {
        var dRow = debriefVals[i];
        var dId = safeGet(dRow, hmDebrief, '要請番号');
        var dStatus = String(safeGet(dRow, hmDebrief, 'ステータス')).trim();
        if (dId && dStatus === '完了') completedIds[dId] = true;
      }

      var caseVals = caseSheet.getDataRange().getValues();
      var now = new Date();
      var cutoff = new Date(now.getTime() - 60 * 24 * 60 * 60 * 1000);
      var pending = [];
      for (var i = caseVals.length - 1; i >= 1; i--) {
        var cRow = caseVals[i];
        var cId = safeGet(cRow, hmCase, '要請番号');
        if (!cId || completedIds[cId]) continue;
        var savedAtRaw = safeGet(cRow, hmCase, '保存日時');
        var savedAt = savedAtRaw ? new Date(String(savedAtRaw).replace(/\//g, '-')) : null;
        if (savedAt && !isNaN(savedAt.getTime()) && savedAt < cutoff) continue;
        var keywords = safeGet(cRow, hmCase, 'キーワード');
        var summary = safeGet(cRow, hmCase, '概要・経過');
        var label = keywords || (summary ? String(summary).slice(0, 20) : cId);
        pending.push({ id: cId, label: label, date: safeGet(cRow, hmCase, '搬入日') });
        if (pending.length >= 5) break;
      }
      pending.forEach(function (p) {
        result.alerts.push({ id: p.id, message: "「" + p.label + "」事案の振り返りシートが未入力です。", link: "debriefing.html?id=" + encodeURIComponent(p.id) });
      });
    }
  } catch (e) {
    Logger.log("未完了アラート集計エラー: " + e.toString());
  }

  return result;
}

// ----------------------------------------------------
// 【全マスタ完全統合取得ロジック】
// ----------------------------------------------------
function getMasterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("スプレッドシートに接続できません。");
  var data = {
    drugs: [], equipment: [], scenarios: [], staff: [], weightMaster: [], pewsMaster: [], treatments: [],
    keywords: [], hospitals: {}, emsList: {}, ventilators: [], settings: {}, tests: []
  };

  try {
    // 1. マスタ管理シート (全角/半角自動判定)
    var sheetAdmin = ss.getSheetByName('マスタ＿管理') || ss.getSheetByName('マスタ_管理');
    if (sheetAdmin) {
      var adminData = sheetAdmin.getDataRange().getValues();
      for (var i = 1; i < adminData.length; i++) {
        var r = adminData[i];
        if (r.length > 1 && r[1] != null && r[1] !== "") { var kw = String(r[1]).trim(); if (kw !== "") data.keywords.push(kw); }
        if (r.length > 3 && r[2] != null && r[3] != null) {
          var hCat = String(r[2]).trim(); var hName = String(r[3]).trim();
          if (hCat !== "" && hName !== "") { if (!data.hospitals[hCat]) data.hospitals[hCat] = []; data.hospitals[hCat].push(hName); }
        }
        if (r.length > 5 && r[4] != null && r[5] != null) {
          var eReg = String(r[4]).trim(); var eName = String(r[5]).trim();
          if (eReg !== "" && eName !== "") { if (!data.emsList[eReg]) data.emsList[eReg] = []; data.emsList[eReg].push(eName); }
        }
        if (r.length > 10 && r[6] != null && r[7] != null) {
          var maker = String(r[6]).trim(); var model = String(r[7]).trim();
          if (maker !== "" && model !== "") {
            data.ventilators.push({ maker: maker, model: model, minW: parseFloat(r[8]) || 0, maxW: parseFloat(r[9]) || 999, desc: String(r[10] || "").trim() });
          }
        }
        if (r.length > 12 && r[11] != null && r[12] != null) {
          var sKey = String(r[11]).trim(); var sVal = String(r[12]).trim();
          if (sKey !== "") data.settings[sKey] = isNaN(parseFloat(sVal)) ? sVal : parseFloat(sVal);
        }
      }
    }

    // デフォルト補定
    if (Object.keys(data.settings).length === 0) {
      data.settings = { Defib_First_J_per_kg: 2, Defib_Second_J_per_kg: 4, Defib_Max_J: 200, Defib_Pad_Adult_Weight: 10, Defib_Steps: "10,15,20,30,50,70,100,150,200", Tube_Depth_Age_Coef: 0.5, Tube_Depth_Age_Base: 12, Tube_Depth_Wt_Coef: 0.5, Tube_Depth_Wt_Base: 8, Adrenaline_Dose_per_kg: 0.01, Adrenaline_Max_Dose: 10 };
    }

    // 2. スタッフマスタ
    var sheetStaff = ss.getSheetByName('マスタ＿スタッフ') || ss.getSheetByName('マスタースタッフ') || ss.getSheetByName('マスタ_スタッフ');
    if (sheetStaff) {
      var hmSt = getHeaderMap(sheetStaff); var stVals = sheetStaff.getDataRange().getValues();
      for (var i = 1; i < stVals.length; i++) {
        var r = stVals[i]; if (!r || r.length < 3) continue;
        if (hmSt.hasOwnProperty('フルネーム(選択用)') && r[hmSt['フルネーム(選択用)']]) {
          data.staff.push({ displayName: String(safeGet(r, hmSt, 'フルネーム(選択用)')).trim(), lastName: String(safeGet(r, hmSt, '姓')).trim(), dept: String(safeGet(r, hmSt, '所属')), roleType: String(safeGet(r, hmSt, '職種')).trim(), phone: hmSt.hasOwnProperty('番号') ? String(safeGet(r, hmSt, '番号')).trim() : "" });
        }
      }
    }

    // 3. 薬品単発
    var sheetSingle = ss.getSheetByName('マスタ＿薬品単発');
    if (sheetSingle) {
      var hm = getHeaderMap(sheetSingle); var vals = sheetSingle.getDataRange().getValues();
      for (var i = 1; i < vals.length; i++) {
        var r = vals[i]; if (!r || r.length < 5) continue;
        if (hm.hasOwnProperty('商品名(代表)') && r[hm['商品名(代表)']]) {
          var productDose = (parseFloat(safeGet(r, hm, '製品用量')) || 1) * (String(safeGet(r, hm, '製品単位')).trim().toLowerCase() === 'g' ? 1000 : 1);
          data.drugs.push({ type: "単発薬", category: String(safeGet(r, hm, '区分')), btnName: String(safeGet(r, hm, '表示名') || safeGet(r, hm, '商品名(代表)')).trim(), name: String(safeGet(r, hm, '商品名(代表)')).trim(), spec: safeGet(r, hm, '製品用量') + safeGet(r, hm, '製品単位'), comp: String(safeGet(r, hm, '組成・希釈方法')).trim(), stdDose: parseFloat(safeGet(r, hm, '標準投与量')) || 0, unit: String(safeGet(r, hm, '投与量単位')), concMg: productDose, concVol: parseFloat(safeGet(r, hm, '製品容量(mL)')) || 1, maxDose: parseFloat(safeGet(r, hm, '最大投与量(上限)')) || 9999, note: String(safeGet(r, hm, '備考')) });
        }
      }
    }

    // 4. 薬品持続
    var sheetCont = ss.getSheetByName('マスタ＿薬品持続');
    if (sheetCont) {
      var hmC = getHeaderMap(sheetCont); var valsC = sheetCont.getDataRange().getValues();
      for (var i = 1; i < valsC.length; i++) {
        var r = valsC[i]; if (!r || r.length < 5) continue;
        if (hmC.hasOwnProperty('商品名(代表)') && r[hmC['商品名(代表)']]) {
          var productDoseC = (parseFloat(safeGet(r, hmC, '原液薬量')) || 1) * (String(safeGet(r, hmC, '薬量単位')).trim().toLowerCase() === 'g' ? 1000 : 1);
          var steps = []; for (var j = 1; j <= 10; j++) { var stepKey = 'Step' + j; if (hmC.hasOwnProperty(stepKey) && r[hmC[stepKey]] !== "") steps.push(parseFloat(r[hmC[stepKey]])); }
          data.drugs.push({ type: "持続薬", category: String(safeGet(r, hmC, '区分')), btnName: String(safeGet(r, hmC, '表示名') || safeGet(r, hmC, '商品名(代表)')).trim(), name: String(safeGet(r, hmC, '商品名(代表)')).trim(), comp: String(safeGet(r, hmC, '組成・希釈（表示用）')).trim(), stdDose: 0, unit: String(safeGet(r, hmC, '処方単位') || 'γ'), concMg: productDoseC, concVol: parseFloat(safeGet(r, hmC, '総液量')) || 1, maxDose: 9999, note: String(safeGet(r, hmC, '備考')), steps: steps });
        }
      }
    }

    // 5. 物品マスタ
    var sheetEq = ss.getSheetByName('マスタ＿物品');
    if (sheetEq) {
      var hmE = getHeaderMap(sheetEq); var eVals = sheetEq.getDataRange().getValues();
      for (var i = 1; i < eVals.length; i++) {
        var r = eVals[i]; if (!r || r.length < 3) continue;
        if (hmE.hasOwnProperty('大項目') && r[hmE['大項目']]) {
          // 対象身長・対象年齢はシート上には以前から用意されていたが、
          // これまでgetMasterData()側で読み込んでおらず未使用のまま放置されていた。
          // 体重は測定/推定値として最も信頼できる主指標のまま維持しつつ、
          // 身長・年齢は「体重で選んだサイズが年齢・身長的に妥当か」の
          // クロスチェック用の補助情報としてフロント側に渡す。
          data.equipment.push({ category: String(safeGet(r, hmE, '区分(ABCDE)')), name: String(safeGet(r, hmE, '大項目')).trim(), size: String(safeGet(r, hmE, 'サイズ・規格')) + (safeGet(r, hmE, '単位') ? ' ' + safeGet(r, hmE, '単位') : ''), minW: parseFloat(safeGet(r, hmE, '対象体重_下限(kg)')) || 0, maxW: parseFloat(safeGet(r, hmE, '対象体重_上限(kg)')) || 9999, minH: parseFloat(safeGet(r, hmE, '対象身長_下限(cm)')) || 0, maxH: parseFloat(safeGet(r, hmE, '対象身長_上限(cm)')) || 9999, minAge: parseFloat(safeGet(r, hmE, '対象年齢_下限(歳)')) || 0, maxAge: parseFloat(safeGet(r, hmE, '対象年齢_上限(歳)')) || 999, inStock: String(safeGet(r, hmE, '院内採用(◯/×)')), note: String(safeGet(r, hmE, '備考')) });
        }
      }
    }

    // 6. セット・シナリオマスタ
    var sheetSet = ss.getSheetByName('マスタ＿セット');
    if (sheetSet) {
      var hmS = getHeaderMap(sheetSet); var sVals = sheetSet.getDataRange().getValues();
      // このシートは項目数に上限を設けないため「含まれる項目名」という
      // 同名ヘッダーがD列以降に何十列も並ぶ構造になっている。
      // getHeaderMap()はヘッダー文字列をキーにした単純なマップなので、
      // 同名ヘッダーが複数あると後の列が前の列を上書きしてしまい、
      // hmS['含まれる項目名']は「最初の」列ではなく「最後の」列のインデックスに
      // なっていた。その結果ほぼ全ての行で実際の項目データより後ろから
      // 読み始めることになり、シナリオ/セットの紐づけ項目がほぼ空になっていた
      // （データ自体はシートに残っており消えていない）。
      // ヘッダー行を直接走査して最初の出現列を使うことで回避する。
      var headerRowSet = sVals[0] || [];
      var firstItemCol = -1;
      for (var hc = 0; hc < headerRowSet.length; hc++) { if (String(headerRowSet[hc]).trim() === '含まれる項目名') { firstItemCol = hc; break; } }
      for (var i = 1; i < sVals.length; i++) {
        var r = sVals[i]; if (!r || r.length < 2) continue;
        if (hmS.hasOwnProperty('セット名') && r[hmS['セット名']]) {
          var items = []; if (firstItemCol !== -1) { for (var j = firstItemCol; j < r.length; j++) { if (r[j]) items.push(String(r[j]).trim()); } }
          data.scenarios.push({ name: String(r[hmS['セット名']]).trim(), type: String(safeGet(r, hmS, '区分')), items: items });
        }
      }
    }

    // 7. 体格マスタ
    var sheetWeight = ss.getSheetByName('マスタ＿体格');
    if (sheetWeight) {
      var hmW = getHeaderMap(sheetWeight); var wtVals = sheetWeight.getDataRange().getValues();
      for (var i = 1; i < wtVals.length; i++) {
        var r = wtVals[i]; if (!r || r.length < 10) continue;
        if (hmW.hasOwnProperty('月齢') && r[hmW['月齢']] !== "") {
          data.weightMaster.push({ months: parseInt(safeGet(r, hmW, '月齢')) || 0, m_low: parseFloat(safeGet(r, hmW, '男子_体重3%')) || 0, m_mid: parseFloat(safeGet(r, hmW, '男子_体重50%')) || 0, m_high: parseFloat(safeGet(r, hmW, '男子_体重97%')) || 0, m_h_low: parseFloat(safeGet(r, hmW, '男子_身長3%（下限）')) || 0, m_h_mid: parseFloat(safeGet(r, hmW, '男子_身長50%')) || 0, m_h_high: parseFloat(safeGet(r, hmW, '男子_身長97%（上限）')) || 0, f_low: parseFloat(safeGet(r, hmW, '女子_体重3%')) || 0, f_mid: parseFloat(safeGet(r, hmW, '女子_体重50%')) || 0, f_high: parseFloat(safeGet(r, hmW, '女子_体重97%')) || 0, f_h_low: parseFloat(safeGet(r, hmW, '女子_身長3%（下限）')) || 0, f_h_mid: parseFloat(safeGet(r, hmW, '女子_身長50%')) || 0, f_h_high: parseFloat(safeGet(r, hmW, '女子_身長97%（上限）')) || 0 });
        }
      }
    }

    // 8. PEWS基準値マスタ
    var sheetPews = ss.getSheetByName('マスタ＿小児基準値');
    if (sheetPews) {
      var hmP = getHeaderMap(sheetPews); var pVals = sheetPews.getDataRange().getValues();
      for (var i = 1; i < pVals.length; i++) {
        var r = pVals[i]; if (!r || r.length < 5) continue;
        if (hmP.hasOwnProperty('判定項目') && r[hmP['判定項目']]) {
          data.pewsMaster.push({ item: String(safeGet(r, hmP, '判定項目')).trim(), ageMin: parseInt(safeGet(r, hmP, '年齢下限')) || 0, ageMax: parseInt(safeGet(r, hmP, '年齢上限')) || 99, normalMin: parseInt(safeGet(r, hmP, '正常値下限')) || 0, normalMax: parseInt(safeGet(r, hmP, '正常値上限')) || 999, th1: parseInt(safeGet(r, hmP, 'スコア1しきい値')) || 0, th2: parseInt(safeGet(r, hmP, 'スコア2しきい値')) || 999, th3: parseInt(safeGet(r, hmP, 'スコア3しきい値')) || 999 });
        }
      }
    }

    // 9. 処置マスタ
    var sheetTx = ss.getSheetByName('マスタ＿処置');
    if (sheetTx) {
      var hmTx = getHeaderMap(sheetTx); var txVals = sheetTx.getDataRange().getValues();
      var cKbn = ""; var cDai = ""; var cChu = "";
      for (var i = 1; i < txVals.length; i++) {
        var r = txVals[i]; if (!r || r.length === 0) continue;
        if (safeGet(r, hmTx, '区分') !== "") cKbn = String(safeGet(r, hmTx, '区分')).trim();
        if (safeGet(r, hmTx, '大項目') !== "") { cDai = String(safeGet(r, hmTx, '大項目')).trim(); cChu = ""; }
        if (safeGet(r, hmTx, '中項目') !== "") cChu = String(safeGet(r, hmTx, '中項目')).trim();
        var opt = hmTx.hasOwnProperty('選択肢') && r[hmTx['選択肢']] != null ? String(r[hmTx['選択肢']]).trim() : "";
        if (opt !== "") data.treatments.push({ category: cKbn, large: cDai, medium: cChu, name: opt, style: String(safeGet(r, hmTx, '方式')).trim() });
      }
    }

    // 10. 検査マスタ
    var sheetTest = ss.getSheetByName('マスタ＿検査');
    if (sheetTest) {
      var tData = sheetTest.getDataRange().getValues();
      var hMapTest = {}; var dataStartRow = 1;
      for (var i = 0; i < tData.length; i++) {
        if (tData[i].indexOf('大項目') !== -1 || tData[i].indexOf('検査項目名') !== -1) {
          for (var c = 0; c < tData[i].length; c++) { if (tData[i][c]) hMapTest[String(tData[i][c]).trim()] = c; }
          dataStartRow = i + 1; break;
        }
      }
      for (var i = dataStartRow; i < tData.length; i++) {
        var r = tData[i]; var name = safeGet(r, hMapTest, '検査項目名');
        if (name) {
          data.tests.push({ category: safeGet(r, hMapTest, '大項目'), subCategory: safeGet(r, hMapTest, '検体種別'), name: String(name).trim(), min: safeGet(r, hMapTest, '正常値下限'), max: safeGet(r, hMapTest, '正常値上限'), unit: safeGet(r, hMapTest, '単位'), note: safeGet(r, hMapTest, '備考') });
        }
      }
    }

  } catch (e) { throw new Error("マスタ一括読み込みエラー: " + e.toString()); }

  // 作戦盤・デブリーフィングが要求する全オブジェクトを絶対に漏らさず返却
  return {
    status: "success",
    masters: { staffList: data.staff, keywordList: data.keywords, hospitalList: data.hospitals, emsList: data.emsList },
    drugs: data.drugs,
    equipment: data.equipment,
    scenarios: data.scenarios,
    weightMaster: data.weightMaster,
    pewsMaster: data.pewsMaster,
    treatments: data.treatments,
    tests: data.tests,
    ventilators: data.ventilators,
    settings: data.settings
  };
}
