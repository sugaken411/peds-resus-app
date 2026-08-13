// バージョン: V6.23 (ポータル連携: お知らせ・マニュアル・未完了アラート追加)
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
      var hm = getHeaderMap(dbSheet);
      var timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
      var p = parsedPayload.payload || {}; var pd = p.data || {}; var roles = p.staffRoles || {};
      var dataDict = {
        '分野区分': '小児', '保存日時': timestamp, '要請番号': pd["要請番号"] || '', '搬入日': pd["日付"] || '', '事案種別': pd["事案種別"] || '', '紹介元医療機関': pd["紹介元"] || '', '年齢/月齢': p.ageText || '', '性別': pd["性別"] || '', '計算体重': p.weight || '', '目安身長': p.height || '', 'キーワード': (p.keywords || []).join(', '), '概要・経過': pd["概要"] || '', '搬入前処置': pd["搬入前処置"] || '', 'AMPL': pd["AMPL"] || '', 'PAT': pd["PAT"] || '', 'バイタル': pd["バイタル"] || '', 'PEWS': pd["PEWS"] || '', '統括': (roles.leader || []).join(', '), '気道管理': (roles.airway || []).join(', '), '胸骨圧迫': (roles.cpr || []).join(', '), 'ルート・薬剤': (roles.route || []).join(', '), '記録': (roles.record || []).join(', '), 'その他役割': (roles.other || []).join(', '), '想定シナリオ・薬剤物品詳細': pd["プロトコル詳細"] || ''
      };
      dbSheet.appendRow(createRowData(hm, dataDict));
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. 事案アーカイブ検索
    if (parsedPayload.action === "fetch_history") {
      var dbSheet = ss.getSheetByName('事案アーカイブ');
      if (!dbSheet) throw new Error("事案アーカイブが見つかりません");
      var hm = getHeaderMap(dbSheet); var data = dbSheet.getDataRange().getValues(); var results = [];
      var sMonth = parsedPayload.month || ""; var sStaff = parsedPayload.staff || ""; var sKw = parsedPayload.keyword || "";

      for (var i = 1; i < data.length; i++) {
        var row = data[i]; var id = safeGet(row, hm, '要請番号'); var date = safeGet(row, hm, '搬入日'); var summary = safeGet(row, hm, '概要・経過'); var keywords = safeGet(row, hm, 'キーワード');
        if (sKw && id.indexOf(sKw) === -1 && summary.indexOf(sKw) === -1 && keywords.indexOf(sKw) === -1) continue;
        if (sMonth && String(date).indexOf(sMonth) === -1) continue;
        results.push({
          id: id, date: date, age: safeGet(row, hm, '年齢/月齢'), weight: safeGet(row, hm, '計算体重'), sex: safeGet(row, hm, '性別'), summary: summary, keywords: keywords, preTx: safeGet(row, hm, '搬入前処置'), protocol: safeGet(row, hm, '想定シナリオ・薬剤物品詳細'), vitals: safeGet(row, hm, 'バイタル'), pat: safeGet(row, hm, 'PAT'), caseType: safeGet(row, hm, '事案種別'), facility: safeGet(row, hm, '紹介元医療機関'),
          roles: { leader: safeGet(row, hm, '統括'), airway: safeGet(row, hm, '気道管理'), cpr: safeGet(row, hm, '胸骨圧迫'), route: safeGet(row, hm, 'ルート・薬剤'), record: safeGet(row, hm, '記録') }
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
          debriefData = { actualDiff: safeGet(data[i], hm, '想定との相違点'), gapBad: safeGet(data[i], hm, 'ギャップ課題'), gapGood: safeGet(data[i], hm, 'ギャップ良かった点'), gapMaster: safeGet(data[i], hm, 'マスタ改修提案'), team: safeGet(data[i], hm, 'チーム連携評価'), action: safeGet(data[i], hm, '次回アクションプラン'), dDate: safeGet(data[i], hm, '開催日'), dTime: safeGet(data[i], hm, '開始時間'), dEndTime: safeGet(data[i], hm, '終了時間'), dPlace: safeGet(data[i], hm, '場所'), timeline: safeGet(data[i], hm, 'タイムラインJSON') }; break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", debriefData: debriefData })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6. デブリーフィング保存
    if (parsedPayload.action === "submit_debrief") {
      var dbSheet = ss.getSheetByName('振り返りアーカイブ');
      if (!dbSheet) throw new Error("振り返りアーカイブが見つかりません");
      var hm = getHeaderMap(dbSheet); var p = parsedPayload.payload || {};
      var dataDict = {
        'タイムスタンプ': Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"), '要請番号': p.id || '', '発生日': p.date || '', '患者情報': p.patientInfo || '', '参加スタッフ': p.debriefStaffs || '', '事案概要': p.summary || '', 'タイムラインJSON': JSON.stringify(p.timeline || []), '記録者': p.recorder || '', '実際の対応スタッフ': p.actualStaffs || '', '想定との相違点': p.actualDiff || '', 'ギャップ課題': p.gapBad || '', 'ギャップ良かった点': p.gapGood || '', 'チーム連携評価': p.team || '', 'マスタ改修提案': p.gapMaster || '', '次回アクションプラン': p.action || '', '開催日': p.dDate || '', '開始時間': p.dTime || '', '終了時間': p.dEndTime || '', '所要時間': p.dDuration || '', '場所': p.dPlace || '', 'ステータス': p.status || ''
      };
      var data = dbSheet.getDataRange().getValues(); var targetRow = -1;
      for (var i = 1; i < data.length; i++) { if (safeGet(data[i], hm, '要請番号') === p.id) { targetRow = i + 1; break; } }
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
      var promptText = "あなたは医療記録アシスタントです。提供されたデータから指定JSONのみ出力。マークダウン不可。\n";
      if (target === "timeline") promptText += "【抽出】時間(time)と内容(contentText)の配列。catは'処置','薬剤','記事'。例:[{\"time\":\"14:23\",\"cat\":\"記事\",\"contentText\":\"挿管\"}]";
      else promptText += "【抽出】" + target + "に関する数値をJSON形式で出力。";

      var parts = [{ text: promptText }];
      if (parsedPayload.imageText) parts.push({ text: "【テキスト】\n" + parsedPayload.imageText });
      if (parsedPayload.imageBase64) parts.push({ inlineData: { mimeType: parsedPayload.mimeType || "image/jpeg", data: parsedPayload.imageBase64.replace(/^data:image\/[a-z]+;base64,/, "") } });

      var res = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey, {
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
      date: safeGet(r, hmNews, '日付'),
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
          data.equipment.push({ category: String(safeGet(r, hmE, '区分(ABCDE)')), name: String(safeGet(r, hmE, '大項目')).trim(), size: String(safeGet(r, hmE, 'サイズ・規格')) + (safeGet(r, hmE, '単位') ? ' ' + safeGet(r, hmE, '単位') : ''), minW: parseFloat(safeGet(r, hmE, '対象体重_下限(kg)')) || 0, maxW: parseFloat(safeGet(r, hmE, '対象体重_上限(kg)')) || 9999, note: String(safeGet(r, hmE, '備考')) });
        }
      }
    }

    // 6. セット・シナリオマスタ
    var sheetSet = ss.getSheetByName('マスタ＿セット');
    if (sheetSet) {
      var hmS = getHeaderMap(sheetSet); var sVals = sheetSet.getDataRange().getValues();
      for (var i = 1; i < sVals.length; i++) {
        var r = sVals[i]; if (!r || r.length < 2) continue;
        if (hmS.hasOwnProperty('セット名') && r[hmS['セット名']]) {
          var items = []; if (hmS.hasOwnProperty('含まれる項目名')) { for (var j = hmS['含まれる項目名']; j < r.length; j++) { if (r[j]) items.push(String(r[j]).trim()); } }
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
