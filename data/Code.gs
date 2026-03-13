const SHEET_RAW       = "폼 응답 1";
const SHEET_SUMMARY   = "집계";
const SHEET_DASHBOARD = "대시보드";

const TYPES  = ["리딩사기", "코인사기", "부동산사기", "기타"];
const ROUTES = ["홈페이지", "전화", "SNS", "지인소개"];

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SHEET_SUMMARY);
  }
  summarySheet.clearContents();

  summarySheet.getRange("A1").setValue("[ 상담 유형별 ]");
  summarySheet.getRange("A1").setFontWeight("bold");
  TYPES.forEach(function(type, i) {
    summarySheet.getRange(i + 2, 1).setValue(type);
    summarySheet.getRange(i + 2, 2).setValue(0);
  });

  summarySheet.getRange("D1").setValue("[ 접수 채널별 ]");
  summarySheet.getRange("D1").setFontWeight("bold");
  ROUTES.forEach(function(route, i) {
    summarySheet.getRange(i + 2, 4).setValue(route);
    summarySheet.getRange(i + 2, 5).setValue(0);
  });

  summarySheet.getRange("G1").setValue("[ 월별 집계 ]");
  summarySheet.getRange("G1").setFontWeight("bold");
  var months = ["1월","2월","3월","4월","5월","6월",
                "7월","8월","9월","10월","11월","12월"];
  months.forEach(function(m, i) {
    summarySheet.getRange(i + 2, 7).setValue(m);
    summarySheet.getRange(i + 2, 8).setValue(0);
  });

  var dashSheet = ss.getSheetByName(SHEET_DASHBOARD);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(SHEET_DASHBOARD);
  }
  dashSheet.clearContents();
  dashSheet.getRange("A1").setValue("대시보드");
}

function onFormSubmit(e) {
  var responses = e.values;
  var consultType = responses[1];
  var route       = responses[2];
  var timestamp   = new Date(responses[0]);
  var month       = timestamp.getMonth() + 1;

  updateTypeCount(consultType);
  updateRouteCount(route);
  updateMonthCount(month);
}

function updateTypeCount(type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(SHEET_SUMMARY);
  var idx = TYPES.indexOf(type);
  if (idx === -1) return;
  var cell = sheet.getRange(idx + 2, 2);
  cell.setValue(cell.getValue() + 1);
}

function updateRouteCount(route) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(SHEET_SUMMARY);
  var idx = ROUTES.indexOf(route);
  if (idx === -1) return;
  var cell = sheet.getRange(idx + 2, 5);
  cell.setValue(cell.getValue() + 1);
}

function updateMonthCount(month) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(SHEET_SUMMARY);
  var cell = sheet.getRange(month + 1, 8);
  cell.setValue(cell.getValue() + 1);
}

function testRun() {
  var fakeEvent = {
    values: [
      new Date().toString(),
      "리딩사기",
      "홈페이지",
      "서울",
      "테스트"
    ]
  };
  onFormSubmit(fakeEvent);
}

function resetCounts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(SHEET_SUMMARY);
  TYPES.forEach(function(_, i)  { sheet.getRange(i + 2, 2).setValue(0); });
  ROUTES.forEach(function(_, i) { sheet.getRange(i + 2, 5).setValue(0); });
  for (var m = 1; m <= 12; m++) {
    sheet.getRange(m + 1, 8).setValue(0);
  }
}

function createForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 폼 생성
  var form = FormApp.create("법률 상담 문의 접수");

  // 질문 1: 상담 유형
  var typeItem = form.addMultipleChoiceItem();
  typeItem.setTitle("상담 유형")
          .setChoiceValues(["리딩사기", "코인사기", "부동산사기", "기타"])
          .setRequired(true);

  // 질문 2: 접수 경로
  var routeItem = form.addMultipleChoiceItem();
  routeItem.setTitle("접수 경로")
           .setChoiceValues(["홈페이지", "전화", "SNS", "지인소개"])
           .setRequired(true);

  // 질문 3: 지역
  var regionItem = form.addTextItem();
  regionItem.setTitle("지역")
            .setRequired(false);

  // 질문 4: 메모
  var memoItem = form.addParagraphTextItem();
  memoItem.setTitle("메모 (선택)")
          .setRequired(false);

  // 폼 응답을 스프레드시트에 연결
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log("✅ 폼 생성 완료: " + form.getEditUrl());
}
// ================================================
// 7. 가짜 데이터 자동 생성 (300건)
// ================================================
function insertFakeData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName("Form_Responses");
  if (!rawSheet) rawSheet = ss.getSheets()[0];

  // 기존 데이터 삭제
  var lastRow = rawSheet.getLastRow();
  if (lastRow > 1) rawSheet.deleteRows(2, lastRow - 1);

  var types   = ["리딩사기", "코인사기", "부동산사기", "기타"];
  var routes  = ["홈페이지", "전화", "SNS", "지인소개"];
  var regions = ["서울", "부산", "인천", "대구", "경기", "광주", "대전", "울산"];
  var typeWeights  = [45, 30, 15, 10];
  var routeWeights = [35, 30, 20, 15];

  function weightedRandom(arr, weights) {
    var total = weights.reduce(function(a, b) { return a + b; }, 0);
    var r = Math.random() * total;
    var cumulative = 0;
    for (var i = 0; i < arr.length; i++) {
      cumulative += weights[i];
      if (r <= cumulative) return arr[i];
    }
    return arr[arr.length - 1];
  }

  // ✅ 메모리에서 집계 (시트 접근 없음)
  var typeCounts  = {"리딩사기":0, "코인사기":0, "부동산사기":0, "기타":0};
  var routeCounts = {"홈페이지":0, "전화":0, "SNS":0, "지인소개":0};
  var monthCounts = {};
  for (var m = 1; m <= 12; m++) monthCounts[m] = 0;

  var rows = [];
  for (var i = 0; i < 300; i++) {
    var t = weightedRandom(types, typeWeights);
    var r = weightedRandom(routes, routeWeights);
    var region = regions[Math.floor(Math.random() * regions.length)];
    var randomMonth = Math.floor(Math.random() * 15);
    var fakeDate = new Date(2025, randomMonth, Math.floor(Math.random() * 28) + 1);
    var month = fakeDate.getMonth() + 1;

    rows.push([fakeDate, t, r, region, ""]);

    typeCounts[t]++;
    routeCounts[r]++;
    if (month >= 1 && month <= 12) monthCounts[month]++;
  }

  // ✅ 폼 응답 시트 한 번에 쓰기
  rawSheet.getRange(2, 1, rows.length, 5).setValues(rows);

  // ✅ 집계 시트 한 번에 쓰기
  var summarySheet = ss.getSheetByName(SHEET_SUMMARY);

  var typeValues = types.map(function(t) { return [typeCounts[t]]; });
  summarySheet.getRange(2, 2, types.length, 1).setValues(typeValues);

  var routeValues = routes.map(function(r) { return [routeCounts[r]]; });
  summarySheet.getRange(2, 5, routes.length, 1).setValues(routeValues);

  var monthValues = [];
  for (var m = 1; m <= 12; m++) monthValues.push([monthCounts[m]]);
  summarySheet.getRange(2, 8, 12, 1).setValues(monthValues);

  Logger.log("✅ 300건 생성 완료 (빠른 버전)");
}
// ================================================
// 8. 대시보드 차트 자동 생성
// ================================================
function createDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet  = ss.getSheetByName(SHEET_SUMMARY);
  var dashSheet     = ss.getSheetByName(SHEET_DASHBOARD);

  // 기존 차트 전부 삭제
  var existingCharts = dashSheet.getCharts();
  existingCharts.forEach(function(chart) {
    dashSheet.removeChart(chart);
  });

  // ── 차트 1: 상담 유형별 파이차트 ──────────────
  var typeChart = dashSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(summarySheet.getRange("A1:B5"))
    .setPosition(1, 1, 0, 0)
    .setOption("title", "상담 유형별 분포")
    .setOption("width", 400)
    .setOption("height", 300)
    .build();
  dashSheet.insertChart(typeChart);

  // ── 차트 2: 접수 채널별 막대차트 ──────────────
  var routeChart = dashSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(summarySheet.getRange("D1:E5"))
    .setPosition(1, 7, 0, 0)
    .setOption("title", "접수 채널별 현황")
    .setOption("width", 400)
    .setOption("height", 300)
    .build();
  dashSheet.insertChart(routeChart);

  // ── 차트 3: 월별 추이 라인차트 ────────────────
  var monthChart = dashSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(summarySheet.getRange("G1:H13"))
    .setPosition(18, 1, 0, 0)
    .setOption("title", "월별 상담 접수 추이")
    .setOption("width", 800)
    .setOption("height", 300)
    .build();
  dashSheet.insertChart(monthChart);

  Logger.log("✅ 대시보드 차트 3개 생성 완료");
}