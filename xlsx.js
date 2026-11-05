/* Builds the Master Plan import file. Same layout as Files\ExportHourFromFile.xlsx:
   one sheet, header row, then employee, project, date (numFmt 14), description, hours (numFmt 20). */
(function () {
  const HEADERS = ["מס' עובד", "מספר פרויקט", "תאריך", "תאור", 'סה"כ שעות'];
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG = "http://schemas.openxmlformats.org/package/2006";
  const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml";

  const escapeXml = s => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  /* Excel day number of a yyyy-MM-dd date. */
  function serial(date) {
    const [y, m, d] = date.split("-").map(Number);
    return Date.UTC(y, m - 1, d) / 86400000 + 25569;
  }

  function file(employeeNum, rows) {
    const strings = [];
    const shared = s => {
      let i = strings.indexOf(s);
      if (i < 0) { strings.push(s); i = strings.length - 1; }
      return i;
    };
    const textCell = (ref, s) => `<c r="${ref}" s="0" t="s"><v>${shared(s)}</v></c>`;
    const idCell = (ref, s) => /^[1-9][0-9]*$/.test(s) ? `<c r="${ref}"><v>${s}</v></c>` : textCell(ref, s);

    let sheetRows = '<row r="1">' + HEADERS.map((h, i) => textCell("ABCDE"[i] + "1", h)).join("") + "</row>";
    rows.forEach((row, i) => {
      const r = i + 2;
      sheetRows += `<row r="${r}">` + idCell("A" + r, employeeNum) + idCell("B" + r, row.num)
        + `<c r="C${r}" s="1"><v>${serial(row.date)}</v></c>` + textCell("D" + r, row.name)
        + `<c r="E${r}" s="2"><v>${row.hours / 24}</v></c></row>`;
    });

    const sheet = HEAD + `<worksheet xmlns="${MAIN}"><dimension ref="A1:E${rows.length + 1}"/>`
      + '<sheetViews><sheetView rightToLeft="1" tabSelected="1" workbookViewId="0"/></sheetViews>'
      + '<cols><col min="1" max="3" width="12.8" customWidth="1"/><col min="4" max="4" width="29.3" customWidth="1"/><col min="5" max="5" width="19.4" customWidth="1"/></cols>'
      + "<sheetData>" + sheetRows + "</sheetData></worksheet>";
    const sst = HEAD + `<sst xmlns="${MAIN}" count="${strings.length}" uniqueCount="${strings.length}">`
      + strings.map(s => "<si><t>" + escapeXml(s) + "</t></si>").join("") + "</sst>";
    const styles = HEAD + `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>`
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
      + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      + '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
      + '<xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>';
    const workbook = HEAD + `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="שעות" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const workbookRels = HEAD + `<Relationships xmlns="${PKG}/relationships">`
      + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>`
      + `<Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;
    const rootRels = HEAD + `<Relationships xmlns="${PKG}/relationships">`
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
    const types = HEAD + `<Types xmlns="${PKG}/content-types">`
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + `<Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/>`
      + `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CT}.worksheet+xml"/>`
      + `<Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/>`
      + `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}.sharedStrings+xml"/></Types>`;

    return zip([
      ["[Content_Types].xml", types], ["_rels/.rels", rootRels], ["xl/workbook.xml", workbook],
      ["xl/_rels/workbook.xml.rels", workbookRels], ["xl/styles.xml", styles],
      ["xl/sharedStrings.xml", sst], ["xl/worksheets/sheet1.xml", sheet],
    ]);
  }

  const CRC_TABLE = Array.from({ length: 256 }, (_, i) => {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* Minimal ZIP writer: stored entries, no compression, which is all a small .xlsx needs. */
  function zip(files) {
    const utf8 = new TextEncoder();
    const out = [], central = [];
    const le16 = v => [v & 0xff, (v >>> 8) & 0xff];
    const le32 = v => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
    const push = (target, bytes) => { for (const b of bytes) target.push(b); };

    for (const [name, text] of files) {
      const nameBytes = utf8.encode(name), content = utf8.encode(text), offset = out.length;
      const common = [...le16(20), ...le16(0), ...le16(0), ...le16(0), ...le16(0x21),
        ...le32(crc32(content)), ...le32(content.length), ...le32(content.length), ...le16(nameBytes.length), ...le16(0)];
      push(out, le32(0x04034b50)); push(out, common); push(out, nameBytes); push(out, content);
      push(central, le32(0x02014b50)); push(central, le16(20)); push(central, common);
      push(central, le16(0)); push(central, le16(0)); push(central, le16(0)); push(central, le32(0));
      push(central, le32(offset)); push(central, nameBytes);
    }
    const centralOffset = out.length;
    push(out, central);
    push(out, le32(0x06054b50)); push(out, le16(0)); push(out, le16(0));
    push(out, le16(files.length)); push(out, le16(files.length));
    push(out, le32(central.length)); push(out, le32(centralOffset)); push(out, le16(0));
    return new Uint8Array(out);
  }

  window.HoursXlsx = { file };
})();
