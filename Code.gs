/**
 * 스프레드시트 ID 및 시트 이름 설정
 * 제공해주신 URL에서 추출한 스프레드시트 ID를 적용했습니다.
 */
const SPREADSHEET_ID = '1Dn14125p57HY-WyoXkK9tGLWwpMtPlTEDKe0HOwr_tU';
const SHEET_NAME = '공식모음'; // 시트 탭 이름 (없으면 자동 생성됩니다)

/**
 * 웹 앱 접속 시 기본 HTML 페이지를 렌더링합니다.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('나만의 공식 암기 모음집')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 시트를 가져오거나, 해당 이름의 시트가 없으면 새로 생성하고 헤더를 추가합니다.
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 기본 헤더 설정 (A: ID, B: 과목/분류, C: 공식 이름, D: 공식 내용, E: 설명/예시, F: 암기상태)
    const headers = ['ID', '분류', '공식 이름', '공식 (수식/코드)', '설명 및 팁', '암기 여부'];
    sheet.appendRow(headers);
    
    // 헤더 스타일링
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
               .setBackground('#4f46e5')
               .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    
    // 샘플 데이터 1개 추가
    sheet.appendRow([
      Utilities.getUuid(),
      '수학',
      '근의 공식',
      '$$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$',
      '이차방정식 ax² + bx + c = 0 (a ≠ 0)의 해를 구하는 공식',
      '미완료'
    ]);
  }
  return sheet;
}

/**
 * 스프레드시트에서 모든 공식 데이터를 읽어옵니다.
 */
function getFormulas() {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    // 2행부터 마지막 행까지 데이터 가져오기
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    return data.map(row => ({
      id: String(row[0]),
      category: String(row[1]),
      title: String(row[2]),
      formula: String(row[3]),
      description: String(row[4]),
      memorized: row[5] === '완료'
    })).filter(item => item.title.trim() !== ''); // 빈 행 제외
  } catch (error) {
    throw new Error('데이터를 불러오는 중 오류가 발생했습니다: ' + error.message);
  }
}

/**
 * 새로운 공식을 스프레드시트에 추가합니다.
 */
function addFormula(formulaData) {
  try {
    const sheet = getOrCreateSheet();
    const id = Utilities.getUuid();
    
    sheet.appendRow([
      id,
      formulaData.category || '기타',
      formulaData.title,
      formulaData.formula,
      formulaData.description || '',
      '미완료'
    ]);
    
    return { success: true, id: id };
  } catch (error) {
    throw new Error('공식 저장 중 오류가 발생했습니다: ' + error.message);
  }
}

/**
 * 특정 공식의 암기 상태를 토글(완료 <-> 미완료)합니다.
 */
function toggleMemorizedStatus(id, isMemorized) {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false };
    
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        const statusText = isMemorized ? '완료' : '미완료';
        sheet.getRange(i + 2, 6).setValue(statusText);
        return { success: true };
      }
    }
    return { success: false, message: '해당 ID를 찾을 수 없습니다.' };
  } catch (error) {
    throw new Error('상태 업데이트 중 오류가 발생했습니다: ' + error.message);
  }
}

/**
 * 특정 공식을 스프레드시트에서 삭제합니다.
 */
function deleteFormula(id) {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false };
    
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, message: '삭제할 항목을 찾을 수 없습니다.' };
  } catch (error) {
    throw new Error('삭제 중 오류가 발생했습니다: ' + error.message);
  }
}
