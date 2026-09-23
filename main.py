function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // 웹뷰어에서 날아온 구버전 저장 요청 방어 (체크박스 상태 포함)
    if (data && data.action === "saveLayout") {
      var personName = data.targetName || "전체";
      PropertiesService.getScriptProperties().setProperty("COLUMN_LAYOUT_" + personName, JSON.stringify(data.layout));
      if (data.isShrinkFit !== undefined) {
        PropertiesService.getScriptProperties().setProperty("SHRINK_FIT_" + personName, data.isShrinkFit ? "Y" : "N");
      }
      return ContentService.createTextOutput(JSON.stringify({result: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    if (!Array.isArray(data) || data.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({result: "empty", added: 0})).setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var baseSheet = ss.getSheetByName("기본");
    if (!baseSheet) baseSheet = ss.insertSheet("기본");
    baseSheet.clearContents();
    baseSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    var processedResult = processAndFormatClaimData();
    
    var listSheet = ss.getSheetByName("발송자리스트");
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

function doGet(e) {
  try {
    var targetName = (e && e.parameter && e.parameter.name) ? String(e.parameter.name).trim() : "전체";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var baseSheet = ss.getSheetByName("기본");
    if (!baseSheet) baseSheet = ss.getSheetByName("클레임"); 
    
    var rawData = baseSheet.getDataRange().getDisplayValues();
    var rawHeaders = rawData.length > 0 ? rawData[0].map(function(h) { return String(h).trim(); }) : [];
    
    var statusColIndex = rawHeaders.indexOf("접수여부");
    var progressColIndex = rawHeaders.indexOf("진행상태");
    var data = [];
    if(rawData.length > 0) data.push(rawHeaders);
    
    for(var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var rawStatus = statusColIndex !== -1 ? String(row[statusColIndex]).replace(/\s+/g, "").toUpperCase() : "";
      var rawProgress = progressColIndex !== -1 ? String(row[progressColIndex]).trim() : "";
      var isProgressMatched = (progressColIndex === -1) || (rawProgress === "확인");
      if (rawStatus !== "Y" && isProgressMatched) {
        data.push(row);
      }
    }

    var defaultWidths = {"품목담당자명": 80, "반품번호": 130, "등록일시": 85, "거래처코드": 75, "거래처명": 150, "반품전화번호": 110, "품목코드": 90, "품목명": 220, "요청내역": 350, "수량": 45, "확인자명": 80, "확인일시": 85};

    var savedLayoutRaw = PropertiesService.getScriptProperties().getProperty("COLUMN_LAYOUT_" + targetName);
    var savedLayout = savedLayoutRaw ? JSON.parse(savedLayoutRaw) : null;
    
    // 💡 [새 기능] 폰트 축소 체크박스 상태 불러오기
    var shrinkFitRaw = PropertiesService.getScriptProperties().getProperty("SHRINK_FIT_" + targetName);
    var isShrinkFit = (shrinkFitRaw === "Y"); 

    if (!savedLayout || !Array.isArray(savedLayout)) {
      var targetColumns = ["품목담당자명", "반품번호", "등록일시", "거래처코드", "거래처명", "반품전화번호", "품목코드", "품목명", "요청내역", "수량", "확인자명", "확인일시"];
      savedLayout = rawHeaders.map(function(h) { 
        return { name: h, visible: targetColumns.indexOf(h) !== -1, width: defaultWidths[h] || 100 }; 
      });
      savedLayout.sort(function(a, b) {
        var idxA = targetColumns.indexOf(a.name);
        var idxB = targetColumns.indexOf(b.name);
        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
        if (idxA !== -1) return -1;
        if (idxB !== -1) return 1;
        return 0;
      });
    } else {
      var existingNames = savedLayout.map(function(item) { return item.name.trim(); });
      rawHeaders.forEach(function(h) {
        if (existingNames.indexOf(h) === -1) {
          savedLayout.push({ name: h, visible: false, width: defaultWidths[h] || 100 });
        }
      });
      savedLayout.forEach(function(item) {
        item.name = item.name.trim();
        if (!item.width) item.width = defaultWidths[item.name] || 100;
      });
    }

    var displayCols = [];
    savedLayout.forEach(function(item) {
      if (item.visible) {
        var idx = rawHeaders.indexOf(item.name);
        if (idx !== -1) {
          displayCols.push({ idx: idx, name: item.name, width: item.width });
        }
      }
    });

    var html = "<style>" +
      ".resizer { width: 6px; height: 100%; position: absolute; right: 0; top: 0; cursor: col-resize; user-select: none; z-index: 10; transition: background-color 0.2s; }" +
      ".resizer:hover { background-color: #3498db; }" +
      "th { position: relative; transition: opacity 0.2s; }" +
      ".drag-handle { cursor: grab; padding-bottom: 5px; user-select: none; }" +
      ".drag-handle:active { cursor: grabbing; }" +
      "th.over { border: 2px dashed #e74c3c !important; opacity: 0.7; }" +
      "body { margin: 0; background-color: #f8f9fa; }" +
      "td { box-sizing: border-box; overflow: hidden; }" +
      ".cell-content { font-size: 13px; line-height: 1.4; transition: font-size 0.1s; }" +
      /* 💡 폰트 축소 활성화 시 줄바꿈 방지 */
      "#claimTable.shrink-mode .cell-content { white-space: nowrap; overflow: hidden; }" +
      /* 비활성화 시 자동 줄바꿈 허용 */
      "#claimTable:not(.shrink-mode) .cell-content { white-space: normal; word-break: break-word; }" +
      "</style>";

    html += "<div style='font-family: \"Malgun Gothic\", sans-serif; padding: 20px; width: 95%; margin: 0 auto; background: #fff; box-shadow: 0 0 10px rgba(0,0,0,0.1); border-radius: 8px; margin-top: 20px;'>";
    
    html += "<div style='display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #3498db; padding-bottom: 12px; margin-bottom: 20px;'>";
    html += "<h2 style='color: #2c3e50; margin: 0;'>📊 " + targetName + " 담당자 클레임 내역</h2>";
    
    // 💡 [새 기능] 상단 버튼 옆에 체크박스 탑재
    html += "<div style='display: flex; gap: 15px; align-items: center;'>" +
            "<label style='font-size: 13px; font-weight: bold; cursor: pointer; display: flex; align-items: center; gap: 5px; color: #2c3e50;'>" +
            "<input type='checkbox' id='shrinkCheck' onchange='toggleShrinkToFit()' " + (isShrinkFit ? "checked" : "") + " style='transform: scale(1.2); cursor: pointer;'>" +
            "셀크기에 맞춰 폰트수정</label>" +
            "<button onclick='openModal()' style='padding: 8px 14px; font-size: 13px; background-color: #7f8c8d; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;'>👁️ 열 꺼내기 / 숨기기</button>" +
            "<button onclick='saveLayoutToServer()' id='saveBtnMain' style='padding: 8px 14px; font-size: 13px; background-color: #2980b9; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;'>💾 현재 서식(위치/크기) 저장</button>" +
            "</div></div>";

    if (data.length > 1) {
      html += "<div style='overflow-x: auto; width: 100%; border-radius: 8px; border: 1px solid #eee;'>";
      // 💡 [핵심] width: 0; 을 설정하여 억지로 100% 채워지는 현상 차단! 지정한 크기 그대로 고정됩니다.
      html += "<table id='claimTable' class='" + (isShrinkFit ? "shrink-mode" : "") + "' style='border-collapse: collapse; width: 0; font-size: 13px; text-align: center; table-layout: fixed;' border='1'>";
      html += "<thead style='background-color: #f2f2f2; font-weight: bold;'><tr>";

      displayCols.forEach(function(col, index) {
        html += "<th draggable='true' data-col-name='" + col.name + "' style='padding: 10px; border: 1px solid #ddd; vertical-align: top; width: " + col.width + "px; min-width: " + col.width + "px; max-width: " + col.width + "px; word-break: break-word;'>";
        html += "<div class='drag-handle' title='마우스로 꾹 눌러서 끌면 순서가 바뀝니다.'>" + col.name + "</div>";
        html += "<input type='text' class='col-filter' data-idx='" + index + "' onkeyup='filterTableMulti()' placeholder='검색' style='width: 100%; box-sizing: border-box; padding: 5px; font-size: 11px; border: 1px solid #ccc; border-radius: 4px; text-align: center; outline: none;'>";
        html += "<div class='resizer' title='드래그하여 크기 조절'></div>";
        html += "</th>";
      });
      html += "</tr></thead><tbody id='tableBody'>";

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
            var cellValue = (data[i] && data[i][col.idx] !== undefined) ? data[i][col.idx] : ""; 
            
            // 💡 폰트 축소 엔진을 작동시키기 위해 텍스트를 div(cell-content)로 한 번 감싸줍니다.
            html += "<td style='padding: " + padding + "; border: 1px solid #ddd; text-align: " + textAlign + "; width: " + col.width + "px; min-width: " + col.width + "px; max-width: " + col.width + "px;'>" +
                    "<div class='cell-content' style='width: 100%;'>" + cellValue + "</div></td>";
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
    
    html += "<div id='layoutModal' style='display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:9999; align-items:center; justify-content:center;'>" +
      "<div style='background:#fff; padding:20px 25px; border-radius:10px; width:380px; max-height:85vh; display:flex; flex-direction:column; box-shadow:0 4px 15px rgba(0,0,0,0.3);'>" +
      
      "<div style='display:flex; justify-content:space-between; align-items:center; border-bottom:2px solid #3498db; padding-bottom:10px; margin-bottom:15px;'>" +
      "<h3 style='margin:0; color:#2c3e50;'>👁️ 추가 열 꺼내기/숨기기</h3>" +
      "<button onclick='saveLayoutToServer(true)' id='saveBtnModal' style='padding:6px 12px; font-size:12px; background-color:#e74c3c; color:#fff; border:none; border-radius:5px; cursor:pointer; font-weight:bold;'>💾 닫고 적용하기</button>" +
      "</div>" +
      
      "<p style='font-size:12px; color:#666; margin-bottom:15px; line-height:1.4;'>체크박스를 클릭하여 원하는 열을 화면에 표시하세요.<br/>설정 후 반드시 우측 상단의 <b>[적용하기]</b>를 눌러주세요.</p>" +
      
      "<div style='overflow-y:auto; flex-grow:1; margin-bottom:15px; border:1px solid #eee; border-radius:5px; padding:10px;'>" +
      "<ul id='layoutList' style='list-style:none; padding:0; margin:0;'></ul>" +
      "</div>" +
      
      "<div style='text-align:right; border-top:1px solid #eee; padding-top:15px;'>" +
      "<button onclick='closeModal()' style='padding:8px 15px; border:1px solid #ccc; background:#f5f5f5; border-radius:5px; cursor:pointer;'>취소 (닫기)</button>" +
      "</div>" +
      "</div></div>";

    html += "<script>" +
      "var layout = " + JSON.stringify(savedLayout) + ";" +
      "var targetName = '" + targetName + "';" +
      
      "function openModal() { renderLayoutList(); document.getElementById('layoutModal').style.display = 'flex'; }" +
      "function closeModal() { document.getElementById('layoutModal').style.display = 'none'; }" +
      
      "function renderLayoutList() {" +
      "  var ul = document.getElementById('layoutList'); ul.innerHTML = '';" +
      "  layout.forEach(function(item, idx) {" +
      "    var li = document.createElement('li');" +
      "    li.style.cssText = 'display:flex; align-items:center; padding:8px 5px; border-bottom:1px solid #f9f9f9;';" +
      "    li.innerHTML = \"<input type='checkbox' id='chk_\" + idx + \"' \" + (item.visible ? 'checked' : '') + \" onchange='layout[\" + idx + \"].visible = this.checked' style='cursor:pointer; transform:scale(1.2); margin-right:10px;'> \" +" +
      "                   \"<label for='chk_\" + idx + \"' style='font-weight:bold; font-size:13px; cursor:pointer; color:#333; width:100%;'>\" + item.name + \"</label>\";" +
      "    ul.appendChild(li);" +
      "  });" +
      "}" +

      "function updateLayoutFromDOM() {" +
      "  var ths = document.querySelectorAll('#claimTable th');" +
      "  var newLayout = [];" +
      "  var seenNames = {};" +
      "  ths.forEach(function(th) {" +
      "    var colName = th.getAttribute('data-col-name');" +
      "    if (colName) {" +
      "      colName = colName.trim();" +
      "      var w = th.offsetWidth || parseInt(th.style.width) || 100;" +
      "      newLayout.push({ name: colName, visible: true, width: w });" +
      "      seenNames[colName] = true;" +
      "    }" +
      "  });" +
      "  layout.forEach(function(item) {" +
      "    var trimmedName = item.name.trim();" +
      "    if (!seenNames[trimmedName]) {" +
      "      newLayout.push({ name: trimmedName, visible: item.visible, width: item.width || 100 });" +
      "      seenNames[trimmedName] = true;" +
      "    }" +
      "  });" +
      "  layout = newLayout;" +
      "}" +
      
      // 💡 [핵심] 저장 버튼을 누를 때 폰트축소 체크박스 상태(isShrinkFit)도 함께 서버에 전송합니다!
      "function saveLayoutToServer(isFromModal) {" +
      "  if (!isFromModal) updateLayoutFromDOM();" +
      "  var btn = isFromModal ? document.getElementById('saveBtnModal') : document.getElementById('saveBtnMain');" +
      "  var originalText = btn.innerText;" +
      "  btn.innerText = '⏳ 저장 중...'; btn.disabled = true;" +
      "  var isShrinkFit = document.getElementById('shrinkCheck').checked;" +
      "  var layoutStr = JSON.stringify(layout);" +
      
      "  if (typeof google !== 'undefined' && google.script && google.script.run) {" +
      "    google.script.run" +
      "      .withSuccessHandler(function(res) {" +
      "        location.reload();" +
      "      })" +
      "      .withFailureHandler(function(err) {" +
      "        alert('❌ 통신 오류: 다시 시도해주세요.');" +
      "        btn.innerText = originalText; btn.disabled = false;" +
      "      })" +
      "      .saveUserLayoutApp(targetName, layoutStr, isShrinkFit);" +
      "  } else {" +
      "    alert('⚠️ 구글 서버에 연결할 수 없습니다.');" +
      "    btn.innerText = originalText; btn.disabled = false;" +
      "  }" +
      "}" +

      // 💡 폰트 축소 체크박스 토글 함수
      "function toggleShrinkToFit() {" +
      "  var isShrink = document.getElementById('shrinkCheck').checked;" +
      "  var table = document.getElementById('claimTable');" +
      "  if (isShrink) {" +
      "    table.classList.add('shrink-mode');" +
      "    adjustFontSizes();" +
      "  } else {" +
      "    table.classList.remove('shrink-mode');" +
      "    var cells = document.querySelectorAll('#claimTable tbody td .cell-content');" +
      "    cells.forEach(function(c) { c.style.fontSize = '13px'; });" +
      "  }" +
      "}" +

      // 💡 엑셀의 "셀에 맞춤" 폰트 자동 축소 엔진!
      "function adjustFontSizes() {" +
      "  if (!document.getElementById('shrinkCheck').checked) return;" +
      "  var tds = document.querySelectorAll('#claimTable tbody td');" +
      "  tds.forEach(function(td) {" +
      "    var content = td.querySelector('.cell-content');" +
      "    if (!content) return;" +
      "    content.style.fontSize = '13px';" + 
      "    content.style.display = 'inline-block';" + // 실제 텍스트 길이를 재기 위함
      "    var available = td.clientWidth - 16;" + // 패딩 영역 제외
      "    var actual = content.scrollWidth;" +
      "    if (actual > available && available > 0) {" +
      "       var scale = available / actual;" +
      "       var newSize = Math.max(8, Math.floor(13 * scale));" + // 최소 폰트는 8px로 제한
      "       content.style.fontSize = newSize + 'px';" +
      "    }" +
      "    content.style.display = 'block';" +
      "  });" +
      "}" +

      "function filterTableMulti() {" +
      "  var tbody = document.getElementById('tableBody');" +
      "  if (!tbody) return;" +
      "  var tr = tbody.getElementsByTagName('tr');" +
      "  var inputs = document.getElementsByClassName('col-filter');" +
      "  for (var i = 0; i < tr.length; i++) {" +
      "    var tdList = tr[i].getElementsByTagName('td');" +
      "    var showRow = true;" +
      "    for (var j = 0; j < inputs.length; j++) {" +
      "      var filterVal = inputs[j].value.toLowerCase().trim();" +
      "      var colIdx = inputs[j].getAttribute('data-idx');" +
      "      if (filterVal !== '') {" +
      "        var cellText = tdList[colIdx] ? tdList[colIdx].textContent.toLowerCase() : '';" +
      "        if (cellText.indexOf(filterVal) === -1) {" +
      "          showRow = false; break;" +
      "        }" +
      "      }" +
      "    }" +
      "    tr[i].style.display = showRow ? '' : 'none';" +
      "  }" +
      "}" +

      "var startX, startWidth, currentTh;" +
      "window.onload = function() {" +
      "  var resizers = document.querySelectorAll('.resizer');" +
      "  for (var i = 0; i < resizers.length; i++) {" +
      "    resizers[i].addEventListener('mousedown', function(e) {" +
      "      currentTh = this.parentElement;" +
      "      startX = e.pageX;" +
      "      startWidth = currentTh.offsetWidth;" +
      "      document.addEventListener('mousemove', mouseMoveHandler);" +
      "      document.addEventListener('mouseup', mouseUpHandler);" +
      "      this.style.backgroundColor = '#2980b9';" +
      "      e.stopPropagation(); e.preventDefault();" +
      "    });" +
      "  }" +

      "  var ths = document.querySelectorAll('#claimTable th');" +
      "  var dragSrcIdx = -1;" +
      "  for (var k = 0; k < ths.length; k++) {" +
      "    ths[k].addEventListener('dragstart', function(e) {" +
      "      if (e.target.tagName === 'INPUT' || e.target.className.indexOf('resizer') > -1) { e.preventDefault(); return; }" +
      "      dragSrcIdx = Array.prototype.indexOf.call(this.parentNode.children, this);" +
      "      e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/html', this.innerHTML);" +
      "      this.style.opacity = '0.4';" +
      "    });" +
      "    ths[k].addEventListener('dragover', function(e) { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; return false; });" +
      "    ths[k].addEventListener('dragenter', function(e) { this.classList.add('over'); });" +
      "    ths[k].addEventListener('dragleave', function(e) { this.classList.remove('over'); });" +
      "    ths[k].addEventListener('drop', function(e) {" +
      "      e.stopPropagation(); this.classList.remove('over');" +
      "      var dropTargetIdx = Array.prototype.indexOf.call(this.parentNode.children, this);" +
      "      if (dragSrcIdx !== dropTargetIdx && dragSrcIdx !== -1) {" +
      "        var table = document.getElementById('claimTable');" +
      "        for (var r = 0; r < table.rows.length; r++) {" +
      "          var moving = table.rows[r].children[dragSrcIdx];" +
      "          var target = table.rows[r].children[dropTargetIdx];" +
      "          if (dragSrcIdx < dropTargetIdx) { target.parentNode.insertBefore(moving, target.nextSibling); } " +
      "          else { target.parentNode.insertBefore(moving, target); }" +
      "        }" +
      "        updateLayoutFromDOM();" +
      "        var inputs = document.getElementsByClassName('col-filter');" +
      "        for (var f = 0; f < inputs.length; f++) {" +
      "          var pTh = inputs[f].closest('th');" +
      "          inputs[f].setAttribute('data-idx', Array.prototype.indexOf.call(pTh.parentNode.children, pTh));" +
      "        }" +
      "      }" +
      "      return false;" +
      "    });" +
      "    ths[k].addEventListener('dragend', function(e) {" +
      "      this.style.opacity = '1';" +
      "      var allTh = document.querySelectorAll('#claimTable th');" +
      "      for(var j=0; j<allTh.length; j++) allTh[j].classList.remove('over');" +
      "    });" +
      "  }" +
      
      // 💡 페이지 로드 완료 시 체크박스 상태에 따라 폰트 축소 적용
      "  if (document.getElementById('shrinkCheck').checked) {" +
      "    adjustFontSizes();" +
      "  }" +
      "};" +

      "function mouseMoveHandler(e) {" +
      "  if (currentTh) {" +
      "    var newWidth = startWidth + (e.pageX - startX);" +
      "    if (newWidth > 15) {" + // 💡 최소 너비를 50px에서 15px로 대폭 완화!
      "      currentTh.style.width = newWidth + 'px';" +
      "      currentTh.style.minWidth = newWidth + 'px';" +
      "      currentTh.style.maxWidth = newWidth + 'px';" +
      "    }" +
      "  }" +
      "}" +
      "function mouseUpHandler(e) {" +
      "  if (currentTh) {" +
      "    currentTh.querySelector('.resizer').style.backgroundColor = 'transparent';" +
      "    updateLayoutFromDOM();" +
      "    currentTh = null;" +
      "    document.removeEventListener('mousemove', mouseMoveHandler);" +
      "    document.removeEventListener('mouseup', mouseUpHandler);" +
      "    adjustFontSizes();" + // 💡 크기 조절을 마쳤을 때 즉시 폰트 축소/확대 다시 계산
      "  }" +
      "}" +
      "</script>";

    html += "</div>";
    return HtmlService.createHtmlOutput(html).setTitle(targetName + " 뷰어");

  } catch (err) {
    return HtmlService.createHtmlOutput("<h3 style='color:red;'>❌ 오류 발생:</h3><p>" + err.toString() + "</p>");
  }
}

// 💡 [핵심] 체크박스 상태(isShrinkFit)도 서버에서 받아서 저장 처리!
function saveUserLayoutApp(personName, layoutStr, isShrinkFit) {
  try {
    PropertiesService.getScriptProperties().setProperty("COLUMN_LAYOUT_" + personName, layoutStr);
    PropertiesService.getScriptProperties().setProperty("SHRINK_FIT_" + personName, isShrinkFit ? "Y" : "N");
    return "success";
  } catch(e) {
    throw new Error(e.toString());
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
  var baseSheet = ss.getSheetByName("기본");
  var claimSheet = ss.getSheetByName("클레임");
  
  if (!baseSheet) return {};

  var targetColumns = ["품목담당자명", "반품번호", "등록일시", "거래처코드", "거래처명", "반품전화번호", "품목코드", "품목명", "요청내역", "수량", "확인자명", "확인일시"];
  var columnWidths = {"품목담당자명": 80, "반품번호": 150, "등록일시": 75, "거래처코드": 70, "거래처명": 200, "반품전화번호": 100, "품목코드": 100, "품목명": 230, "요청내역": 400, "수량": 50, "확인자명": 150, "확인일시": 75};

  var claimValues = baseSheet.getDataRange().getValues();
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

  var listSheet = ss.getSheetByName("발송자리스트");
  var lastRowList = listSheet ? listSheet.getLastRow() : 0;
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

  var rawUrl = ScriptApp.getService().getUrl();
  var webAppUrl = rawUrl ? rawUrl.replace(/\/u\/\d+\//, "/") : "";

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
  Logger.log("✅ 반응형 폰트 축소 엔진 및 최소크기 해제 메일 발송 완료");
}
