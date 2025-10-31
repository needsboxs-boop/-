function setup_NEEDSBOXES() {
  const ss = SpreadsheetApp.getActive();

  const cfg = {
    products: {
      name: 'Products_All',
      priceNum: 'Price_SAR_numeric',
      priceLbl: 'Price_SAR_label',
      proposed: 'Proposed_Price_SAR',
      multiplier: 1.25,
    },
    assets: {
      name: 'Assets_Index',
      base: 'Base_URL',
      url: 'URL',
      src: 'UTM_Source',
      med: 'UTM_Medium',
      camp: 'UTM_Campaign',
      cont: 'UTM_Content',
      auto: 'UTM_URL (auto)',
    },
  };

  function byHeader(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    const lastRow = sh.getLastRow() || 1;
    const lastCol = sh.getLastColumn() || 1;
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const map = {};
    headers.forEach((h, i) => (map[h] = i + 1));
    sh.setFrozenRows(1);
    sh.getRange(1, 1, Math.max(1, lastRow), lastCol).setWrap(true);
    sh.autoResizeColumns(1, lastCol);
    return { sh, lastRow, lastCol, map };
  }

  // Products_All: تنسيق الأسعار + سعر مقترح
  try {
    const { sh, lastRow, map } = byHeader(cfg.products.name);
    if (lastRow > 1) {
      const r2 = 2;
      const n = lastRow - 1;
      const cNum = map[cfg.products.priceNum] || 0;
      const cLbl = map[cfg.products.priceLbl] || 0;
      const cProp = map[cfg.products.proposed] || 0;

      if (cNum) sh.getRange(r2, cNum, n, 1).setNumberFormat('#,##0.00');
      if (cLbl) sh.getRange(r2, cLbl, n, 1).setNumberFormat('#,##0 "ر.س"');

      if (cProp && cNum) {
        sh
          .getRange(r2, cProp, n, 1)
          .setFormulaR1C1(`=R[0]C[${cNum - cProp}]*${cfg.products.multiplier}`)
          .setNumberFormat('#,##0.00');
      }
    }
  } catch (e) {}

  // Assets_Index: توليد رابط UTM تلقائي
  try {
    const { sh, lastRow, map, lastCol } = byHeader(cfg.assets.name);
    if (lastRow > 1) {
      const r2 = 2;
      const n = lastRow - 1;
      const cBase = map[cfg.assets.base] || 0;
      const cSrc = map[cfg.assets.src] || 0;
      const cMed = map[cfg.assets.med] || 0;
      const cCamp = map[cfg.assets.camp] || 0;
      const cCont = map[cfg.assets.cont] || 0;
      const cAuto = map[cfg.assets.auto] || 0;

      if (cAuto && cBase && cSrc && cMed && cCamp) {
        // إذا وُجد ? استخدم & وإلا استخدم ?
        const f = `
=IF(LEN(INDIRECT(ADDRESS(ROW(),${cBase})))=0,"",
INDIRECT(ADDRESS(ROW(),${cBase})) &
IF(REGEXMATCH(INDIRECT(ADDRESS(ROW(),${cBase})),"\\?"),"&", "?") &
"utm_source=" & INDIRECT(ADDRESS(ROW(),${cSrc})) &
"&utm_medium=" & INDIRECT(ADDRESS(ROW(),${cMed})) &
"&utm_campaign=" & INDIRECT(ADDRESS(ROW(),${cCamp})) &
IF(LEN(INDIRECT(ADDRESS(ROW(),${cCont})))>0, "&utm_content=" & INDIRECT(ADDRESS(ROW(),${cCont})), "")
)`;
        sh.getRange(r2, cAuto, n, 1).setFormula(f.trim());
        sh.autoResizeColumns(1, lastCol);
      }
    }
  } catch (e) {}
}
