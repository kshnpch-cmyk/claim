function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // 💡 1. 사용자가 웹페이지에서 '컬럼 서식 저장'을 눌렀을 때 '해당 담당자 이름'으로 따로 저장!
    if (data && data.action === "saveLayout") {
      var personName = data.targetName || "전체";
      PropertiesService.getScriptProperties().setProperty("COLUMN_LAYOUT_" + personName, JSON.stringify(data.layout));
      return ContentService.createTextOutput(JSON.stringify({result: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 💡 2. 파이썬(GitHub)에서 들어오는 데이터 처리
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var claimSheet = ss.getSheetByName("클레임") || ss.getSheets()[0];
    var listSheet = ss.getSheetByName("발송자리스트");

    if (!data || data.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({result: "empty", added: 0})).setMimeType(ContentService.MimeType.JSON);
    }

    claimSheet.clearContents();
    claimSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    var processedResult = processAndFormatClaimData();
    
    if (listSheet) {
      var lastRow = listSheet.getLastRow();
      if (lastRow > 0) {
        var emailData = listSheet.getRange(1, 2, lastRow, 1).getValues();
        var emailList = [];
        for (var i = 0; i < emailData.length; i++) {
          var emailStr = String(emailData[i][0]).trim();
          if (emailStr.indexOf("@") !== -1) emailList.push(emailStr);
        }
        
        if (emailList.length > 0) {
          var groupEmails = emailList.join(",");
          sendTotalClaimEmail(groupEmails);
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      result: "success", 
      total_rows: data.length - 1,
      sent_persons: Object.keys(processedResult).length
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({result: "error", error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// 💡 [개인별로 저장된 열 서식(순서, 숨김, 너비)을 적용하여 웹페이지 구성]
function doGet(e) {
  try {
    var targetName = (e && e.parameter && e.parameter.name) ? e.parameter.name : "전체";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var claimSheet = ss.getSheetByName("클레임");
    var data = claimSheet.getDataRange().getDisplayValues();

    var webAppUrl = "https://script.google.com/macros/s/AKfycbxACYFrxGPsNimEsbB3rSEUroxm-cHFqmUnL9t2CIhzQGISIZyEirgDXjhHaVOMFvhe/exec";

    // 기본 열 너비 (처음 열었을 때 세팅되는 값)
    var defaultWidths = {"품목담당자명": 80, "반품번호": 130, "등록일시": 85, "거래처코드": 75, "거래처명": 150, "반품전화번호": 110, "품목코드": 90, "품목명": 220, "요청내역": 350, "수량": 45, "확인자명": 80, "확인일시": 85};

    // 💡 구글 서버에 저장된 '해당 담당자의 개인 서식' 불러오기
    var savedLayoutRaw = PropertiesService.getScriptProperties().getProperty("COLUMN_LAYOUT_" + targetName);
    var savedLayout = savedLayoutRaw ? JSON.parse(savedLayoutRaw) : null;
    var rawHeaders = data.length > 0 ? data[0] : [];

    if (!savedLayout || !Array.isArray(savedLayout)) {
      savedLayout = rawHeaders.map(function(h) { return { name: h, visible: true, width: defaultWidths[h] || 100 }; });
    } else {
      var existingNames = savedLayout.map(function(item) { return item.name; });
      rawHeaders.forEach(function(h) {
        if (existingNames.indexOf(h) === -1) {
          savedLayout.push({ name: h, visible: true, width: defaultWidths[h] || 100 });
        }
      });
      // 예전 서식에 width가 없을 경우를 대비한 보정
      savedLayout.forEach(function(item) {
        if (!item.width) item.width = defaultWidths[item.name] || 100;
      });
    }

    var displayCols = [];
    savedLayout.forEach(function(item) {
      if (item.visible) {
        var idx = rawHeaders.indexOf(item.name);
        if (idx !== -1) displayCols.push({ idx: idx, name: item.name, width: item.width });
      }
    });

    var html = "<div style='font-family: \"Malgun Gothic\", sans-serif; padding: 20px; max-width: 1500px; margin: 0 auto;'>";
    
    html += "<div style='display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #3498db; padding-bottom: 12px; margin-bottom: 20px;'>";
    html += "<h2 style='color: #2c3e50; margin: 0;'>📊 " + targetName + " 담당자 클레임 내역</h2>";
    
    html += "<div style='display: flex; gap: 10px; align-items: center;'>" +
            "<input type='text' id='filterInput' onkeyup='filterTable()' placeholder='🔍 거래처, 품목명, 내역 검색...' " +
            "style='padding: 8px 14px; font-size: 13px; border: 1px solid #3498db; border-radius: 6px; width: 240px; outline: none;'>" +
            "<button onclick='openModal()' style='padding: 8px 14px; font-size: 13px; background-color: #2c3e50; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;'>⚙️ 내 화면 서식 설정</button>" +
            "</div></div>";

    if (data.length > 1) {
      // 💡 가로 스크롤을 지원하는 컨테이너
      html += "<div style='overflow-x: auto; width: 100%; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.05);'>";
      html += "<table id='claimTable' style='border-collapse: collapse; min-width: 100%; font-size: 13px; text-align: center;' border='1'>";
      html += "<thead style='background-color: #f2f2f2; font-weight: bold;'><tr>";

      displayCols.forEach(function(col) {
        html += "<th style='padding: 12px 8px; border: 1px solid #ddd; min-width: " + col.width + "px; max-width: " + col.width + "px; word-break: break-word;'>" + col.name + "</th>";
      });
      html += "</tr></thead><tbody>";

      var nameIndex = rawHeaders.indexOf("품목담당자명");
      if (nameIndex === -1) nameIndex = 0;

      var count = 0;
      for (var i = 1; i < data.length; i++) {
        var rowName = String(data[i][nameIndex]).trim();
        if (targetName === "전체" || rowName === targetName) {
          html += "<tr style='background-color: " + (count % 2 === 0 ? "#fff" : "#f9f9f9") + ";'>";
          displayCols.forEach(function(col) {
            var alignLeft = (col.name === "요청내역" || col.name === "품목명" || col.name === "거래처명");
            var textAlign = alignLeft ? "left" : "center";
            var padding = alignLeft ? "8px 12px" : "8px";
            
            html += "<td style='padding: " + padding + "; border: 1px solid #ddd; min-width: " + col.width + "px; max-width: " + col.width + "px; word-break: break-word; text-align: " + textAlign + ";'>" + data[i][col.idx] + "</td>";
          });
          html += "</tr>";
          count++;
        }
      }
      html += "</tbody></table></div>";
      
      if (count === 0) {
        html += "<p style='margin-top: 20px; color: #e74c3c; font-weight: bold;'>조회된 클레임 내역이 없습니다.</p>";
      }
    } else {
      html += "<p style='margin-top: 20px;'>데이터가 없습니다.</p>";
    }
    
    // 💡 모달 창 UI (내 전용 서식 저장 안내문구 수정)
    html += "<div id='layoutModal' style='display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:9999; align-items:center; justify-content:center;'>" +
      "<div style='background:#fff; padding:25px; border-radius:10px; width:450px; max-height:85vh; overflow-y:auto; box-shadow:0 4px 15px rgba(0,0,0,0.2);'>" +
      "<h3 style='margin-top:0; border-bottom:2px solid #3498db; padding-bottom:10px; color:#2c3e50;'>⚙️ [" + targetName + "] 전용 서식 설정</h3>" +
      "<p style='font-size:12px; color:#666; margin-bottom:15px; line-height:1.5;'>✔️ 체크 해제 시 열 숨김<br/>✔️ <b>[크기]</b>에 숫자를 입력하여 가로 넓이(px) 변경<br/>✔️ 이 설정은 <b>다른 담당자 화면에는 영향을 주지 않습니다.</b></p>" +
      "<ul id='layoutList' style='list-style:none; padding:0; margin:0 0 20px 0;'></ul>" +
      "<div style='text-align:right; border-top:1px solid #eee; padding-top:15px;'>" +
      "<button onclick='closeModal()' style='padding:8px 15px; margin-right:8px; border:1px solid #ccc; background:#f5f5f5; border-radius:5px; cursor:pointer;'>취소</button>" +
      "<button onclick='saveLayoutToServer()' style='padding:8px 15px; background:#27ae60; color:#fff; border:none; border-radius:5px; cursor:pointer; font-weight:bold;'>💾 내 서식 저장하기</button>" +
      "</div></div></div>";

    // 💡 클라이언트 자바스크립트
    html += "<script>" +
      "var layout = " + JSON.stringify(savedLayout) + ";" +
      "var webAppUrl = '" + webAppUrl + "';" +
      "var targetName = '" + targetName + "';" +  // 현재 뷰어의 이름표 기록
      
      "function openModal() { renderLayoutList(); document.getElementById('layoutModal').style.display = 'flex'; }" +
      "function closeModal() { document.getElementById('layoutModal').style.display = 'none'; }" +
      
      "function renderLayoutList() {" +
      "  var ul = document.getElementById('layoutList'); ul.innerHTML = '';" +
      "  layout.forEach(function(item, idx) {" +
      "    var li = document.createElement('li');" +
      "    li.style.cssText = 'display:flex; align-items:center; justify-content:space-between; padding:8px 0; border-bottom:1px solid #eee;';" +
      
      "    var left = document.createElement('div');" +
      "    left.style.cssText = 'display:flex; align-items:center; gap:6px;';" +
      "    left.innerHTML = \"<input type='checkbox' id='chk_\" + idx + \"' \" + (item.visible ? 'checked' : '') + \" onchange='layout[\" + idx + \"].visible = this.checked' style='cursor:pointer;'> \" +" +
      "                     \"<label for='chk_\" + idx + \"' style='font-weight:bold; font-size:13px; cursor:pointer; min-width:85px; display:inline-block;'>\" + item.name + \"</label> \" +" +
      "                     \"<span style='font-size:11px; color:#888;'>크기:</span> <input type='number' value='\" + item.width + \"' onchange='layout[\" + idx + \"].width = parseInt(this.value) || 100' style='width:45px; padding:3px; font-size:12px; border:1px solid #ccc; border-radius:3px; text-align:right;'> <span style='font-size:11px; color:#888;'>px</span>\";" +
      
      "    var right = document.createElement('div');" +
      "    right.innerHTML = \"<button onclick='moveItem(\" + idx + \", -1)' style='padding:2px 8px; margin-right:3px; cursor:pointer; background:#f0f0f0; border:1px solid #ddd; border-radius:3px;'>▲</button><button onclick='moveItem(\" + idx + \", 1)' style='padding:2px 8px; cursor:pointer; background:#f0f0f0; border:1px solid #ddd; border-radius:3px;'>▼</button>\";" +
      
      "    li.appendChild(left); li.appendChild(right); ul.appendChild(li);" +
      "  });" +
      "}" +
      
      "function moveItem(idx, dir) {" +
      "  var targetIdx = idx + dir;" +
      "  if (targetIdx < 0 || targetIdx >= layout.length) return;" +
      "  var temp = layout[idx]; layout[idx] = layout[targetIdx]; layout[targetIdx] = temp;" +
      "  renderLayoutList();" +
      "}" +
      
      "function saveLayoutToServer() {" +
      "  if(!confirm('[" + targetName + "] 담당자님의 전용 서식을 이대로 저장하시겠습니까?\\n(다른 사람의 화면에는 영향을 주지 않습니다.)')) return;" +
      "  fetch(webAppUrl, {" +
      "    method: 'POST'," +
      "    headers: {'Content-Type': 'text/plain'}," +
      "    body: JSON.stringify({ action: 'saveLayout', layout: layout, targetName: targetName })" +
      "  })" +
      "  .then(function(res){ return res.json(); })" +
      "  .then(function(data){" +
      "    if(data.result === 'success') {" +
      "      alert('✅ 내 전용 서식이 성공적으로 저장되었습니다!');" +
      "      location.reload();" +
      "    } else {" +
      "      alert('❌ 저장 실패: ' + JSON.stringify(data));" +
      "    }" +
      "  })" +
      "  .catch(function(err){ alert('❌ 에러 발생: ' + err); });" +
      "}" +

      "function filterTable() {" +
      "  var input = document.getElementById('filterInput');" +
      "  var filter = input.value.toLowerCase().trim();" +
      "  var table = document.getElementById('claimTable');" +
      "  var tr = table.getElementsByTagName('tr');" +
      "  for (var i = 1; i < tr.length; i++) {" +
      "    var tdList = tr[i].getElementsByTagName('td');" +
      "    var match = false;" +
      "    for (var j = 0; j < tdList.length; j++) {" +
      "      if (tdList[j] && tdList[j].textContent.toLowerCase().indexOf(filter) > -1) {" +
      "        match = true; break;" +
      "      }" +
      "    }" +
      "    tr[i].style.display = match ? '' : 'none';" +
      "  }" +
      "}" +
      "</script>";

    html += "</div>";
    return HtmlService.createHtmlOutput(html).setTitle(targetName + " 뷰어");

  } catch (err) {
    return HtmlService.createHtmlOutput("<h3 style='color:red;'>❌ 오류 발생:</h3><p>" + err.toString() + "</p>");
  }
}

function formatDateToYYYYMMDD(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd");
  var str = String(val).trim();
  var match = str.match(/(\d{4})[-.\/]\s*(\d{1,2})[-.\/]\s*(\d{1,2})/);
  if (match) return match[1] + "-" + ("0" + match[2]).slice(-2) + "-" + ("0" + match[3]).slice(-2);
  return str;
}

function processAndFormatClaimData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listSheet = ss.getSheetByName("발송자리스트");
  var claimSheet = ss.getSheetByName("클레임");

  var targetColumns = ["품목담당자명", "반품번호", "등록일시", "거래처코드", "거래처명", "반품전화번호", "품목코드", "품목명", "요청내역", "수량", "확인자명", "확인일시"];
  var columnWidths = {"품목담당자명": 80, "반품번호": 150, "등록일시": 75, "거래처코드": 70, "거래처명": 200, "반품전화번호": 100, "품목코드": 100, "품목명": 230, "요청내역": 400, "수량": 50, "확인자명": 150, "확인일시": 75};

  var claimValues = claimSheet.getDataRange().getValues();
  if (claimValues.length < 2) return {};

  var originalHeaders = claimValues[0].map(function(h) { return String(h).trim(); });
  var originalRows = claimValues.slice(1);
  var colIndexes = targetColumns.map(function(targetCol) { return originalHeaders.indexOf(targetCol); });
  var statusColIndex = originalHeaders.indexOf("접수여부");
  var progressColIndex = originalHeaders.indexOf("진행상태");
  var regDateIdxInTarget = targetColumns.indexOf("등록일시");
  var confirmDateIdxInTarget = targetColumns.indexOf("확인일시");

  var reorderedRows = originalRows.map(function(row) {
    return colIndexes.map(function(idx, colArrIdx) {
      if (idx !== -1 && idx < row.length) {
        var val = row[idx];
        if (colArrIdx === regDateIdxInTarget || colArrIdx === confirmDateIdxInTarget) return formatDateToYYYYMMDD(val);
        return (val instanceof Date) ? Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd") : String(val).trim();
      }
      return "";
    });
  });

  var lastRowList = listSheet.getLastRow();
  var targetNames = [];
  if (lastRowList > 0) {
    var listData = listSheet.getRange(1, 1, lastRowList, 1).getValues();
    listData.forEach(function(row) {
      var n = String(row[0]).trim();
      if (n && n !== "이름" && n !== "담당자") targetNames.push(n);
    });
  }

  var filteredReorderedRows = reorderedRows.filter(function(r, idx) {
    var origRow = originalRows[idx];
    var targetPerson = String(r[0]).trim();
    var rawStatus = statusColIndex !== -1 ? String(origRow[statusColIndex]).replace(/\s+/g, "").toUpperCase() : "";
    var rawProgress = progressColIndex !== -1 ? String(origRow[progressColIndex]).trim() : "";
    var isProgressMatched = (progressColIndex === -1) || (rawProgress === "확인");
    return targetNames.indexOf(targetPerson) !== -1 && rawStatus !== "Y" && isProgressMatched;
  });

  filteredReorderedRows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0]), 'ko'); });
  claimSheet.clear();

  var finalTableData = [targetColumns].concat(filteredReorderedRows);
  var totalRows = finalTableData.length;
  var range = claimSheet.getRange(1, 1, totalRows, targetColumns.length);
  
  range.setNumberFormat("@");
  range.setValues(finalTableData);

  var maxCols = claimSheet.getMaxColumns();
  if (maxCols > targetColumns.length) claimSheet.deleteColumns(targetColumns.length + 1, maxCols - targetColumns.length);

  var qtyColIdx = targetColumns.indexOf("수량") + 1;
  if (totalRows > 1 && qtyColIdx > 0) claimSheet.getRange(2, qtyColIdx, totalRows - 1, 1).setNumberFormat("#,##0");

  range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  claimSheet.getRange(1, 1, 1, targetColumns.length).setFontWeight("bold").setHorizontalAlignment("center");
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  SpreadsheetApp.flush();

  for (var col = 1; col <= targetColumns.length; col++) {
    var colName = targetColumns[col - 1];
    var specifiedWidth = columnWidths[colName];
    if (specifiedWidth) claimSheet.setColumnWidth(col, specifiedWidth);
  }
  return {};
}

function sendTotalClaimEmail(recipientEmails) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var claimSheet = ss.getSheetByName("클레임");
  var data = claimSheet.getDataRange().getDisplayValues();
  if (data.length < 2) return; 

  var headers = data[0];
  var rows = data.slice(1);
  var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  var subject = "미접수 클레임 전달건 (" + rows.length + "건) - " + todayStr;

  var webAppUrl = "https://script.google.com/macros/s/AKfycbxACYFrxGPsNimEsbB3rSEUroxm-cHFqmUnL9t2CIhzQGISIZyEirgDXjhHaVOMFvhe/exec";

  var nameIndex = headers.indexOf("품목담당자명");
  if (nameIndex === -1) nameIndex = 0;

  var htmlBody = "<div style='font-family: Arial, sans-serif; color: #333;'>" +
    "<h3>안녕하세요. 금일 OMS 미접수 클레임 공유드립니다.</h3>" +
    "<p style='color: #2980b9;'>※ 아래 표에서 <b>담당자 이름(파란색 링크)</b>을 클릭하시면, 해당 담당자의 클레임만 모아서 볼 수 있습니다.</p>" +
    "<table style='border-collapse: collapse; width: 100%; font-size: 12px; text-align: center;' border='1'>" +
    "<thead style='background-color: #f2f2f2; font-weight: bold;'><tr>";

  headers.forEach(function(h) {
    htmlBody += "<th style='padding: 6px; border: 1px solid #ccc;'>" + h + "</th>";
  });
  htmlBody += "</tr></thead><tbody>";

  var seenNames = {};

  rows.forEach(function(row) {
    htmlBody += "<tr>";
    row.forEach(function(cell, colIdx) {
      var cleanCell = String(cell).replace(/^'/, '');
      
      if (colIdx === nameIndex && cleanCell !== "") {
        if (!seenNames[cleanCell]) {
          var linkUrl = webAppUrl + "?name=" + encodeURIComponent(cleanCell);
          htmlBody += "<td style='padding: 5px; border: 1px solid #ccc;'>" +
                      "<a href='" + linkUrl + "' style='color: #0056b3; font-weight: bold; text-decoration: underline;' target='_blank'>" + cleanCell + "</a>" +
                      "</td>";
          seenNames[cleanCell] = true;
        } else {
          htmlBody += "<td style='padding: 5px; border: 1px solid #ccc;'>" + cleanCell + "</td>";
        }
      } else {
        htmlBody += "<td style='padding: 5px; border: 1px solid #ccc;'>" + cleanCell + "</td>";
      }
    });
    htmlBody += "</tr>";
  });

  htmlBody += "</tbody></table><br/>감사합니다.</div>";

  GmailApp.sendEmail(recipientEmails, subject, "", {
    htmlBody: htmlBody
  });
  Logger.log("✅ 개인별 서식 저장 지원 메일 발송 완료");
}
