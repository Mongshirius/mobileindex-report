/**
 * 모바일인덱스 리포트 파이프라인 — Google Apps Script 버전.
 *
 * Python 패키지(report_pipeline)의 포팅.
 *   datasource_ -> transform_ -> buildWorkbook_ -> sendEmail_
 * 오케스트레이터: runPipeline()
 *
 * Python 버전과 달리 SMTP/앱 비밀번호가 필요 없다 — MailApp이 실행 계정 권한으로
 * 직접 발송한다. 데이터 취득부(datasource_)는 더미 스텁이며, 2단계에서
 * UrlFetchApp으로 모바일인덱스 API를 호출하도록 이 함수 내부만 교체한다.
 */

// ─────────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────────

/** "DAILY" = 매일 DAILY_HOUR시 / "TEST_5MIN" = 5분마다 (테스트용). */
var SCHEDULE_MODE = 'DAILY';
var DAILY_HOUR = 8; // REPORT_TZ 기준 시(hour)

var APPS = ['앱A', '앱B', '앱C'];
var LOOKBACK_DAYS = 7;
var METRIC_COLUMNS = ['dau', 'installs', 'revenue'];
var COLUMNS = ['date', 'app_name', 'dau', 'installs', 'revenue'];

/**
 * 스크립트 속성에서 설정을 읽는다 (Python의 환경변수/Secrets에 해당).
 * 프로젝트 설정 > 스크립트 속성 에서 MAIL_TO 를 반드시 추가할 것.
 */
function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  var mailTo = (props.getProperty('MAIL_TO') || '').trim();
  if (!mailTo) {
    throw new Error(
      "스크립트 속성 'MAIL_TO' 가 없습니다. " +
      '프로젝트 설정 > 스크립트 속성에서 추가하세요.'
    );
  }
  var tz = (props.getProperty('REPORT_TZ') || '').trim() ||
           Session.getScriptTimeZone() || 'Asia/Seoul';
  return { mailTo: mailTo, tz: tz };
}

// ─────────────────────────────────────────────────────────────
// 1. datasource_  ← 2단계 교체 지점
// ─────────────────────────────────────────────────────────────

/**
 * 결정적 더미 데이터. runDateStr 기준 최근 LOOKBACK_DAYS일 × APPS.
 * 반환: 헤더를 포함한 2차원 배열 (SpreadsheetApp.setValues 형식).
 *
 * 2단계: 이 함수 몸통을 모바일인덱스 연동으로 교체.
 *   - 경로 A: UrlFetchApp.fetch(API URL, {headers: {Authorization: ...}})
 *   - 경로 B: (Apps Script는 헤드리스 브라우저가 없어 CSV 로그인 자동화 불가)
 * 반환 스키마(COLUMNS)와 dtype(date=ISO 문자열, 지표=숫자)은 유지할 것.
 */
function datasource_(runDateStr) {
  var rows = [COLUMNS.slice()];
  var runDate = new Date(runDateStr + 'T00:00:00Z');
  for (var offset = 0; offset < LOOKBACK_DAYS; offset++) {
    var d = new Date(runDate.getTime() - (LOOKBACK_DAYS - 1 - offset) * 86400000);
    var dStr = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
    var ordinal = Math.floor(d.getTime() / 86400000); // epoch 이후 일수 (결정적)
    for (var ai = 0; ai < APPS.length; ai++) {
      var seed = ordinal + ai * 1000;
      rows.push([
        dStr,
        APPS[ai],
        10000 + (seed % 5000),
        500 + (seed % 300),
        Math.round((seed % 900) * 1.5 * 100) / 100,
      ]);
    }
  }
  return rows;
}

// ─────────────────────────────────────────────────────────────
// 2. transform_
// ─────────────────────────────────────────────────────────────

/**
 * 원본 rows(헤더 포함)를 받아 { sheets, summary } 를 반환.
 *   sheets['원본'] = 원본 그대로
 *   sheets['요약'] = app_name 별 지표 합계
 */
function transform_(rawRows, runDateStr) {
  var header = rawRows[0];
  var body = rawRows.slice(1);

  var missing = [];
  ['date', 'app_name'].concat(METRIC_COLUMNS).forEach(function (c) {
    if (header.indexOf(c) === -1) missing.push(c);
  });
  if (missing.length) {
    throw new Error('입력 데이터에 필수 컬럼이 없습니다: ' + missing.join(', '));
  }

  var idx = {};
  header.forEach(function (c, i) { idx[c] = i; });

  var byApp = {};
  body.forEach(function (r) {
    var app = r[idx.app_name];
    if (!byApp[app]) byApp[app] = { dau: 0, installs: 0, revenue: 0 };
    byApp[app].dau += r[idx.dau];
    byApp[app].installs += r[idx.installs];
    byApp[app].revenue += r[idx.revenue];
  });

  var apps = Object.keys(byApp).sort();
  var summarySheet = [['app_name'].concat(METRIC_COLUMNS)];
  apps.forEach(function (a) {
    summarySheet.push([
      a,
      byApp[a].dau,
      byApp[a].installs,
      Math.round(byApp[a].revenue * 100) / 100,
    ]);
  });

  var dates = body.map(function (r) { return r[idx.date]; }).sort();
  return {
    sheets: { '원본': rawRows, '요약': summarySheet },
    summary: {
      row_count: body.length,
      app_count: apps.length,
      date_range: [dates[0], dates[dates.length - 1]],
      run_date: runDateStr,
      generated_at: new Date().toISOString(),
    },
  };
}

// ─────────────────────────────────────────────────────────────
// 3. buildWorkbook_
// ─────────────────────────────────────────────────────────────

/**
 * data.sheets 의 각 시트를 임시 스프레드시트에 쓰고, 전체를 xlsx blob으로 export.
 * 반환: { blob, tempFileId }.  호출측이 tempFileId를 반드시 휴지통 처리할 것.
 */
function buildWorkbook_(data, runDateStr) {
  var ss = SpreadsheetApp.create('mobileindex_report_tmp_' + Date.now());
  var names = Object.keys(data.sheets); // ['원본', '요약']
  var first = ss.getSheets()[0];
  first.setName(names[0]);

  names.forEach(function (name, i) {
    var sheet = i === 0 ? first : ss.insertSheet(name);
    var rows = data.sheets[name];
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  });
  SpreadsheetApp.flush();

  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
                  '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    throw new Error('xlsx export 실패: HTTP ' + resp.getResponseCode());
  }
  var blob = resp.getBlob().setName('mobileindex_report_' + runDateStr + '.xlsx');
  return { blob: blob, tempFileId: ss.getId() };
}

// ─────────────────────────────────────────────────────────────
// 4. sendEmail_
// ─────────────────────────────────────────────────────────────

function sendEmail_(config, subject, body, blob) {
  MailApp.sendEmail({
    to: config.mailTo,
    subject: subject,
    body: body,
    attachments: [blob],
  });
}

// ─────────────────────────────────────────────────────────────
// 오케스트레이터
// ─────────────────────────────────────────────────────────────

/** 트리거가 호출하는 진입점. 수동 실행도 이 함수. */
function runPipeline() {
  var config = getConfig_();
  var runDateStr = Utilities.formatDate(new Date(), config.tz, 'yyyy-MM-dd');
  console.log('run_date=%s tz=%s 남은 메일 할당량=%s',
              runDateStr, config.tz, MailApp.getRemainingDailyQuota());

  var raw = datasource_(runDateStr);
  var data = transform_(raw, runDateStr);
  var wb = buildWorkbook_(data, runDateStr);

  try {
    var subject = '[모바일인덱스 리포트] ' + runDateStr;
    var s = data.summary;
    var body =
      runDateStr + ' 모바일인덱스 리포트입니다.\n\n' +
      '- 데이터 기간: ' + s.date_range[0] + ' ~ ' + s.date_range[1] + '\n' +
      '- 행 수: ' + s.row_count + '\n' +
      '- 앱 수: ' + s.app_count + '\n' +
      '- 생성 시각: ' + s.generated_at + '\n';
    sendEmail_(config, subject, body, wb.blob);
    console.log('발송 완료 → %s (%s)', config.mailTo, wb.blob.getName());
  } finally {
    DriveApp.getFileById(wb.tempFileId).setTrashed(true);
  }
}

// ─────────────────────────────────────────────────────────────
// 트리거 관리 — 한 번만 실행
// ─────────────────────────────────────────────────────────────

/** 시간 기반 트리거를 등록한다. SCHEDULE_MODE 상수에 따라 매일 / 5분. */
function installTrigger() {
  removeTrigger();
  var config = getConfig_();
  var builder = ScriptApp.newTrigger('runPipeline').timeBased();
  if (SCHEDULE_MODE === 'TEST_5MIN') {
    builder.everyMinutes(5).create();
    console.log('트리거 등록: 5분마다 (테스트)');
  } else {
    builder.atHour(DAILY_HOUR).everyDays(1).inTimezone(config.tz).create();
    console.log('트리거 등록: 매일 %s시 (%s)', DAILY_HOUR, config.tz);
  }
}

/** runPipeline 트리거를 모두 제거한다. */
function removeTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runPipeline') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  console.log('트리거 %s개 제거', removed);
}
