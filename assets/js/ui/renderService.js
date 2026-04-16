(function initForm15RenderService(global) {
  const { htmlEscape } = global.Form15Utils;

  function setStatus(ui, message, type) {
    ui.statusBox.textContent = message;
    ui.statusBox.classList.remove("ok", "err");
    if (type === "ok") ui.statusBox.classList.add("ok");
    if (type === "err") ui.statusBox.classList.add("err");
  }

  function renderTable(ui, rows) {
    if (!rows.length) {
      ui.tableRoot.innerHTML = '<div class="empty">Không tìm thấy dòng nào có "Đánh giá = Test bền".</div>';
      return;
    }
    const dynamicCols = [];
    const seenCols = new Set();
    for (const item of rows) {
      for (const key of Object.keys(item.rowData || {})) {
        if (!seenCols.has(key)) {
          seenCols.add(key);
          dynamicCols.push(key);
        }
      }
    }
    const baseHeaders = ["STT", "Mã tác vụ", "Tên tác vụ", "Sheet", "Link file"];
    const allHeaders = baseHeaders.concat(dynamicCols);

    let html = "<table><thead><tr>";
    for (const h of allHeaders) html += "<th>" + htmlEscape(h) + "</th>";
    html += "</tr></thead><tbody>";

    for (let i = 0; i < rows.length; i += 1) {
      const item = rows[i];
      html += "<tr>";
      html += "<td>" + (i + 1) + "</td>";
      html += "<td>" + htmlEscape(item.taskCode) + "</td>";
      html += "<td>" + htmlEscape(item.taskName) + "</td>";
      html += "<td>" + htmlEscape(item.sheetName) + "<small>Nguồn: " + htmlEscape(item.sourceSheetName) + "</small></td>";
      html += '<td><a href="' + htmlEscape(item.fileUrl) + '" target="_blank" rel="noopener">Mở file</a></td>';
      for (const col of dynamicCols) html += '<td class="mono">' + htmlEscape(item.rowData[col] ?? "") + "</td>";
      html += "</tr>";
    }
    html += "</tbody></table>";
    ui.tableRoot.innerHTML = html;
  }

  global.Form15RenderService = {
    setStatus,
    renderTable,
  };
})(window);

