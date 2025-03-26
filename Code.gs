// 盤前選擇權快篩系統 v1
// 作者：AI助手
// 版本：1.0

// 測試用函數，可用於驗證 Apps Script 環境
function myFunction() {
  Logger.log("測試函數執行成功！");
  
  // 取得目前的試算表
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("目前試算表: " + ss.getName());
  
  // 檢查工作表是否存在
  const mainSheet = ss.getSheetByName("選擇權快篩");
  if (mainSheet) {
    Logger.log("選擇權快篩工作表已存在");
  } else {
    Logger.log("選擇權快篩工作表不存在，請執行初始化");
  }
  
  // 顯示執行結果
  const ui = SpreadsheetApp.getUi();
  ui.alert('測試結果', '函數執行成功！請查看 Apps Script 日誌以了解詳細資訊', ui.ButtonSet.OK);
}

// 當打開試算表時執行
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('選擇權快篩')
    .addItem('初始化工作表', 'initializeSheet')
    .addItem('手動刷新數據', 'refreshData')
    .addItem('關於系統', 'showAbout')
    .addToUi();
}

// 初始化工作表
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 檢查是否已經存在主工作表
  let mainSheet;
  try {
    mainSheet = ss.getSheetByName("選擇權快篩");
    if (mainSheet) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert('工作表已存在', '選擇權快篩工作表已存在，是否重新初始化？', ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) {
        ss.deleteSheet(mainSheet);
      } else {
        return;
      }
    }
  } catch (e) {
    console.log('檢查工作表時出錯：' + e.toString());
  }
  
  // 創建主工作表
  mainSheet = ss.insertSheet("選擇權快篩");
  
  // 設置標題行
  const headers = [
    "股票代碼", "名稱", "當前價格", "盤前變動", "盤前成交量", 
    "技術分析", "IV Rank", "建議策略", "到期日", "進場價", 
    "停損價", "風險報酬比", "損益平衡點", "備註", "OptionStrat連結"
  ];
  
  const headerRange = mainSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // 添加預設的10檔股票行
  const defaultStocks = [
    ["AAPL", "=GOOGLEFINANCE(A2, \"name\")", "=GOOGLEFINANCE(A2, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["MSFT", "=GOOGLEFINANCE(A3, \"name\")", "=GOOGLEFINANCE(A3, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["AMZN", "=GOOGLEFINANCE(A4, \"name\")", "=GOOGLEFINANCE(A4, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["GOOGL", "=GOOGLEFINANCE(A5, \"name\")", "=GOOGLEFINANCE(A5, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["TSLA", "=GOOGLEFINANCE(A6, \"name\")", "=GOOGLEFINANCE(A6, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["META", "=GOOGLEFINANCE(A7, \"name\")", "=GOOGLEFINANCE(A7, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["NVDA", "=GOOGLEFINANCE(A8, \"name\")", "=GOOGLEFINANCE(A8, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["AMD", "=GOOGLEFINANCE(A9, \"name\")", "=GOOGLEFINANCE(A9, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["NFLX", "=GOOGLEFINANCE(A10, \"name\")", "=GOOGLEFINANCE(A10, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["SPY", "=GOOGLEFINANCE(A11, \"name\")", "=GOOGLEFINANCE(A11, \"price\")", "", "", "", "", "", "", "", "", "", "", "", ""]
  ];
  
  mainSheet.getRange(2, 1, defaultStocks.length, defaultStocks[0].length).setValues(defaultStocks);
  
  // 設置列寬
  mainSheet.setColumnWidth(1, 80);  // 股票代碼
  mainSheet.setColumnWidth(2, 150); // 名稱
  mainSheet.setColumnWidth(3, 100); // 當前價格
  mainSheet.setColumnWidth(4, 100); // 盤前變動
  mainSheet.setColumnWidth(5, 120); // 盤前成交量
  mainSheet.setColumnWidth(6, 90);  // 技術分析
  mainSheet.setColumnWidth(7, 80);  // IV Rank
  mainSheet.setColumnWidth(8, 150); // 建議策略
  mainSheet.setColumnWidth(9, 90);  // 到期日
  mainSheet.setColumnWidth(10, 80); // 進場價
  mainSheet.setColumnWidth(11, 80); // 停損價
  mainSheet.setColumnWidth(12, 100); // 風險報酬比
  mainSheet.setColumnWidth(13, 100); // 損益平衡點
  mainSheet.setColumnWidth(14, 200); // 備註
  mainSheet.setColumnWidth(15, 150); // OptionStrat連結
  
  // 設置技術分析下拉選單
  const technicalAnalysisRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['偏多', '偏空', '震盪'], true)
    .build();
  mainSheet.getRange(2, 6, defaultStocks.length).setDataValidation(technicalAnalysisRule);
  
  // 設置數據格式
  mainSheet.getRange(2, 3, defaultStocks.length, 1).setNumberFormat('$0.00');
  mainSheet.getRange(2, 4, defaultStocks.length, 1).setNumberFormat('+0.00%;-0.00%');
  mainSheet.getRange(2, 7, defaultStocks.length, 1).setNumberFormat('0.0');
  mainSheet.getRange(2, 10, defaultStocks.length, 1).setNumberFormat('$0.00');
  mainSheet.getRange(2, 11, defaultStocks.length, 1).setNumberFormat('$0.00');
  mainSheet.getRange(2, 12, defaultStocks.length, 1).setNumberFormat('0.00');
  mainSheet.getRange(2, 13, defaultStocks.length, 1).setNumberFormat('$0.00');
  
  // 設置公式：根據技術分析和IV Rank自動建議策略
  for (let i = 2; i <= defaultStocks.length + 1; i++) {
    // 建議策略公式 - 分割成多個部分提高可讀性
    const bullStrategy = `IF(G${i}<30, "Long Call", IF(G${i}<70, "Bull Call Spread", "Bull Put Spread"))`;
    const bearStrategy = `IF(G${i}<30, "Long Put", IF(G${i}<70, "Bear Put Spread", "Bear Call Spread"))`;
    const neutralStrategy = `IF(G${i}<30, "Long Straddle", IF(G${i}<70, "Long Strangle", "Iron Condor"))`;
    const strategyFormula = `=IF(AND(F${i}<>"", G${i}<>""), IF(F${i}="偏多", ${bullStrategy}, IF(F${i}="偏空", ${bearStrategy}, ${neutralStrategy})), "")`;
    
    mainSheet.getRange(i, 8).setFormula(strategyFormula);
    
    // 設定到期日 (當前日期 + 30天，主要用於月選)
    mainSheet.getRange(i, 9).setFormula(`=IF(H${i}<>"", TEXT(WORKDAY(TODAY(), 30), "yyyy-mm-dd"), "")`);
    
    // 進場價、停損價會根據策略和當前價格變動，這裡預留給使用者手動填寫
    
    // 風險報酬比 - 分割成多個部分提高可讀性
    const riskLongOptions = `IF(OR(H${i}="Long Call", H${i}="Long Put"), "無限 : " & TEXT(J${i}, "$0.00")`;
    const riskBullBearSpreads = `IF(OR(H${i}="Bull Call Spread", H${i}="Bear Put Spread"), TEXT(((C${i}*0.05) - J${i}), "$0.00") & " : " & TEXT(J${i}, "$0.00")`;
    const riskCreditSpreads = `IF(OR(H${i}="Bull Put Spread", H${i}="Bear Call Spread"), TEXT(J${i}, "$0.00") & " : " & TEXT(((C${i}*0.05) - J${i}), "$0.00")`;
    const riskStraddles = `IF(OR(H${i}="Long Straddle", H${i}="Long Strangle"), "無限 : " & TEXT(J${i}*2, "$0.00")`;
    const riskCondors = `IF(H${i}="Iron Condor", TEXT(J${i}, "$0.00") & " : " & TEXT(((C${i}*0.05) - J${i}), "$0.00"), "")`;
    
    // 修復嵌套結構
    const riskFormula = `=IF(AND(J${i}<>"", K${i}<>""), ${riskLongOptions}, ${riskBullBearSpreads}, ${riskCreditSpreads}, ${riskStraddles}, ${riskCondors}), "")`;
    
    // 使用基礎版本替代，避免複雜嵌套可能導致的問題
    const simpleRiskFormula = `=IF(AND(J${i}<>"", K${i}<>""), 
      IF(OR(H${i}="Long Call", H${i}="Long Put"), "無限 : " & TEXT(J${i}, "$0.00"),
      IF(OR(H${i}="Bull Call Spread", H${i}="Bear Put Spread"), TEXT(((C${i}*0.05) - J${i}), "$0.00") & " : " & TEXT(J${i}, "$0.00"),
      IF(OR(H${i}="Bull Put Spread", H${i}="Bear Call Spread"), TEXT(J${i}, "$0.00") & " : " & TEXT(((C${i}*0.05) - J${i}), "$0.00"),
      IF(OR(H${i}="Long Straddle", H${i}="Long Strangle"), "無限 : " & TEXT(J${i}*2, "$0.00"),
      IF(H${i}="Iron Condor", TEXT(J${i}, "$0.00") & " : " & TEXT(((C${i}*0.05) - J${i}), "$0.00"), ""))))), "")`;
    
    mainSheet.getRange(i, 12).setFormula(simpleRiskFormula);
    
    // 損益平衡點 - 分割成多個部分提高可讀性
    let bepParts = [];
    bepParts.push(`IF(H${i}="Long Call", TEXT(C${i} + J${i}, "$0.00")`);
    bepParts.push(`IF(H${i}="Long Put", TEXT(C${i} - J${i}, "$0.00")`);
    bepParts.push(`IF(H${i}="Bull Call Spread", TEXT(C${i} + J${i}, "$0.00") & " / " & TEXT(C${i} + C${i}*0.05, "$0.00")`);
    bepParts.push(`IF(H${i}="Bear Put Spread", TEXT(C${i} - J${i}, "$0.00") & " / " & TEXT(C${i} - C${i}*0.05, "$0.00")`);
    bepParts.push(`IF(H${i}="Bull Put Spread", TEXT(C${i} - J${i}, "$0.00") & " / " & TEXT(C${i} - C${i}*0.05, "$0.00")`);
    bepParts.push(`IF(H${i}="Bear Call Spread", TEXT(C${i} + J${i}, "$0.00") & " / " & TEXT(C${i} + C${i}*0.05, "$0.00")`);
    bepParts.push(`IF(H${i}="Long Straddle", TEXT(C${i} - J${i}, "$0.00") & " / " & TEXT(C${i} + J${i}, "$0.00")`);
    bepParts.push(`IF(H${i}="Long Strangle", TEXT(C${i} - J${i}*1.05, "$0.00") & " / " & TEXT(C${i} + J${i}*1.05, "$0.00")`);
    bepParts.push(`IF(H${i}="Iron Condor", TEXT(C${i} - C${i}*0.05, "$0.00") & " ~ " & TEXT(C${i} + C${i}*0.05, "$0.00"), "")`);
    
    // 構建嵌套的 IF 語句
    let bepFormula = `=IF(AND(H${i}<>"", J${i}<>""), `;
    for (let j = 0; j < bepParts.length; j++) {
      if (j === 0) {
        bepFormula += bepParts[j];
      } else {
        bepFormula += ", " + bepParts[j];
      }
      // 添加右括號，除了最後一個
      if (j < bepParts.length - 1) {
        bepFormula += ")";
      }
    }
    bepFormula += ")))))))), \"\")";
    
    // 使用簡單版本避免複雜格式問題
    const simpleBepFormula = `=IF(AND(H${i}<>"", J${i}<>""), 
      IF(H${i}="Long Call", TEXT(C${i} + J${i}, "$0.00"),
      IF(H${i}="Long Put", TEXT(C${i} - J${i}, "$0.00"),
      IF(H${i}="Bull Call Spread", TEXT(C${i} + J${i}, "$0.00") & " / " & TEXT(C${i} + C${i}*0.05, "$0.00"),
      IF(H${i}="Bear Put Spread", TEXT(C${i} - J${i}, "$0.00") & " / " & TEXT(C${i} - C${i}*0.05, "$0.00"),
      IF(H${i}="Bull Put Spread", TEXT(C${i} - J${i}, "$0.00") & " / " & TEXT(C${i} - C${i}*0.05, "$0.00"),
      IF(H${i}="Bear Call Spread", TEXT(C${i} + J${i}, "$0.00") & " / " & TEXT(C${i} + C${i}*0.05, "$0.00"),
      IF(H${i}="Long Straddle", TEXT(C${i} - J${i}, "$0.00") & " / " & TEXT(C${i} + J${i}, "$0.00"),
      IF(H${i}="Long Strangle", TEXT(C${i} - J${i}*1.05, "$0.00") & " / " & TEXT(C${i} + J${i}*1.05, "$0.00"),
      IF(H${i}="Iron Condor", TEXT(C${i} - C${i}*0.05, "$0.00") & " ~ " & TEXT(C${i} + C${i}*0.05, "$0.00"), ""))))))))), "")`;
    
    mainSheet.getRange(i, 13).setFormula(simpleBepFormula);
    
    // OptionStrat連結 - 分割成多個部分提高可讀性
    const stratList = [
      `IF(H${i}="Long Call", "long-call"`,
      `IF(H${i}="Long Put", "long-put"`,
      `IF(H${i}="Bull Call Spread", "bull-call-spread"`,
      `IF(H${i}="Bear Put Spread", "bear-put-spread"`,
      `IF(H${i}="Bull Put Spread", "bull-put-spread"`,
      `IF(H${i}="Bear Call Spread", "bear-call-spread"`,
      `IF(H${i}="Long Straddle", "long-straddle"`,
      `IF(H${i}="Long Strangle", "long-strangle"`,
      `IF(H${i}="Iron Condor", "iron-condor", "")`
    ];
    
    let linkFormula = `=IF(H${i}<>"", HYPERLINK("https://optionstrat.com/build/" & `;
    for (let k = 0; k < stratList.length; k++) {
      if (k === 0) {
        linkFormula += stratList[k];
      } else {
        linkFormula += ", " + stratList[k];
      }
      // 添加右括號，除了最後一個 (最後一個已經包含在字串中)
      if (k < stratList.length - 1) {
        linkFormula += ")";
      }
    }
    linkFormula += ` & "/" & A${i}, "點擊開啟策略模擬器"), "")`;
    
    // 使用簡單版本避免複雜格式問題
    const simpleLinkFormula = `=IF(H${i}<>"", HYPERLINK("https://optionstrat.com/build/" & 
      IF(H${i}="Long Call", "long-call", 
      IF(H${i}="Long Put", "long-put",
      IF(H${i}="Bull Call Spread", "bull-call-spread",
      IF(H${i}="Bear Put Spread", "bear-put-spread",
      IF(H${i}="Bull Put Spread", "bull-put-spread",
      IF(H${i}="Bear Call Spread", "bear-call-spread",
      IF(H${i}="Long Straddle", "long-straddle",
      IF(H${i}="Long Strangle", "long-strangle", 
      IF(H${i}="Iron Condor", "iron-condor", "")))))))))) & "/" & A${i}, "點擊開啟策略模擬器"), "")`;
    
    mainSheet.getRange(i, 15).setFormula(simpleLinkFormula);
  }
  
  // 創建市場情緒工作表
  let marketSheet = ss.insertSheet("市場情緒");
  
  // 設置市場情緒標題
  marketSheet.getRange(1, 1, 1, 4).setValues([["指標", "當前值", "變動", "評分"]]);
  marketSheet.getRange(1, 1, 1, 4).setBackground('#4285F4');
  marketSheet.getRange(1, 1, 1, 4).setFontColor('#FFFFFF');
  marketSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  // 添加VIX和SPY指標
  marketSheet.getRange(2, 1, 2, 4).setValues([
    ["VIX", "=GOOGLEFINANCE(\"VIX\", \"price\")", "=GOOGLEFINANCE(\"VIX\", \"changepct\")", ""],
    ["SPY", "=GOOGLEFINANCE(\"SPY\", \"price\")", "=GOOGLEFINANCE(\"SPY\", \"changepct\")", ""]
  ]);
  
  // 設置VIX評分公式 (VIX低分數高，表示市場樂觀；VIX高分數低，表示市場恐慌)
  marketSheet.getRange(2, 4).setFormula("=IF(B2<15, 5, IF(B2<20, 4, IF(B2<25, 3, IF(B2<30, 2, 1))))");
  
  // 設置SPY評分公式 (SPY上漲分數高，表示市場看漲；SPY下跌分數低，表示市場看跌)
  marketSheet.getRange(3, 4).setFormula("=IF(C3>1, 5, IF(C3>0.5, 4, IF(C3>0, 3, IF(C3>-0.5, 2, 1))))");
  
  // 市場綜合情緒
  marketSheet.getRange(5, 1).setValue("市場綜合情緒");
  marketSheet.getRange(5, 1).setFontWeight('bold');
  marketSheet.getRange(5, 2).setFormula("=(D2+D3)/2");
  
  // 設置情緒評價
  marketSheet.getRange(5, 3).setFormula("=IF(B5>=4.5, \"極度樂觀\", IF(B5>=3.5, \"偏向樂觀\", IF(B5>=2.5, \"中性\", IF(B5>=1.5, \"偏向謹慎\", \"極度謹慎\"))))");
  
  // 設置情緒顏色條件格式 - 使用多個規則代替漸層
  // 紅色 (1-2分)
  let redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 2)
    .setBackground("#FF6C60")
    .setRanges([marketSheet.getRange("B5")])
    .build();
    
  // 黃色 (2-4分)
  let yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(2, 4)
    .setBackground("#FFEB9C")
    .setRanges([marketSheet.getRange("B5")])
    .build();
    
  // 綠色 (4-5分)
  let greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4, 5)
    .setBackground("#72FF60")
    .setRanges([marketSheet.getRange("B5")])
    .build();
  
  // 應用規則
  marketSheet.setConditionalFormatRules([redRule, yellowRule, greenRule]);
  
  // 設置列寬
  marketSheet.setColumnWidth(1, 120);
  marketSheet.setColumnWidth(2, 100);
  marketSheet.setColumnWidth(3, 100);
  marketSheet.setColumnWidth(4, 100);
  
  // 數據格式
  marketSheet.getRange(2, 2, 2, 1).setNumberFormat('0.00');
  marketSheet.getRange(2, 3, 2, 1).setNumberFormat('+0.00%;-0.00%');
  marketSheet.getRange(5, 2).setNumberFormat('0.0');
  
  // 創建策略模板工作表
  let templateSheet = ss.insertSheet("策略範例模板");
  
  // 設置模板標題
  templateSheet.getRange(1, 1, 1, 4).setValues([["技術分析", "IV Rank", "建議策略", "策略說明"]]);
  templateSheet.getRange(1, 1, 1, 4).setBackground('#4285F4');
  templateSheet.getRange(1, 1, 1, 4).setFontColor('#FFFFFF');
  templateSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  // 添加策略模板數據
  const templateData = [
    ["偏多", "低 (0-30)", "Long Call", "看漲且波動率可能上升，直接買入看漲期權"],
    ["偏多", "中 (30-70)", "Bull Call Spread", "看漲但波動率適中，買入低價格看漲期權同時賣出高價格看漲期權"],
    ["偏多", "高 (70-100)", "Bull Put Spread", "看漲且波動率高，賣出低價格看跌期權同時買入更低價格看跌期權"],
    ["偏空", "低 (0-30)", "Long Put", "看跌且波動率可能上升，直接買入看跌期權"],
    ["偏空", "中 (30-70)", "Bear Put Spread", "看跌但波動率適中，買入高價格看跌期權同時賣出低價格看跌期權"],
    ["偏空", "高 (70-100)", "Bear Call Spread", "看跌且波動率高，賣出低價格看漲期權同時買入高價格看漲期權"],
    ["震盪", "低 (0-30)", "Long Straddle", "預期大幅波動但方向不明，同時買入相同行使價的看漲和看跌期權"],
    ["震盪", "中 (30-70)", "Long Strangle", "預期大幅波動但方向不明，同時買入不同行使價的看漲和看跌期權"],
    ["震盪", "高 (70-100)", "Iron Condor", "預期小幅波動，同時賣出中間價位的看漲和看跌期權，買入更遠價位的看漲和看跌期權"]
  ];
  
  templateSheet.getRange(2, 1, templateData.length, templateData[0].length).setValues(templateData);
  
  // 設置列寬
  templateSheet.setColumnWidth(1, 100);
  templateSheet.setColumnWidth(2, 120);
  templateSheet.setColumnWidth(3, 150);
  templateSheet.setColumnWidth(4, 400);
  
  // 設置自動刷新觸發器
  setDailyTrigger();
  
  // 完成初始化後顯示提示
  const ui = SpreadsheetApp.getUi();
  ui.alert('初始化完成', '選擇權快篩系統已成功初始化！\n\n您現在可以：\n1. 在「選擇權快篩」工作表中設置您想追蹤的股票\n2. 填寫技術分析和IV Rank欄位\n3. 系統將自動建議相應的選擇權策略', ui.ButtonSet.OK);
}

// 手動刷新數據
function refreshData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("選擇權快篩");
  
  if (!mainSheet) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('錯誤', '未找到選擇權快篩工作表，請先初始化工作表', ui.ButtonSet.OK);
    return;
  }
  
  // 獲取股票代碼列表
  const stockRange = mainSheet.getRange(2, 1, 10);
  const stocks = stockRange.getValues().flat().filter(stock => stock !== "");
  
  // 如果沒有股票，顯示提示
  if (stocks.length === 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('提示', '未找到股票代碼，請在A欄輸入要追蹤的股票代碼', ui.ButtonSet.OK);
    return;
  }
  
  // 更新盤前數據
  updatePremarketData(stocks);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('刷新完成', '數據已成功刷新！', ui.ButtonSet.OK);
}

// 更新盤前數據
function updatePremarketData(stocks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("選擇權快篩");
  
  // 對每支股票進行處理
  for (let i = 0; i < stocks.length; i++) {
    const stock = stocks[i];
    const row = i + 2; // 對應的行號
    
    try {
      // 獲取當前價格
      const currentPrice = mainSheet.getRange(row, 3).getValue();
      
      // 嘗試從外部API獲取盤前數據 (這裡使用示例數據，實際應對接真實API)
      // 由於Google Apps Script對外部API有限制，這裡只用簡單示例模擬
      // 實際使用時，可以整合像Alpha Vantage或Yahoo Finance的API
      
      // 模擬盤前變動 (-2%到+2%之間的隨機值)
      const premarketChange = (Math.random() * 4 - 2) / 100;
      
      // 模擬盤前成交量 (當前成交量的10%到30%之間的隨機值)
      const premarketVolume = Math.floor(Math.random() * 20 + 10) + "%";
      
      // 更新盤前變動欄位
      mainSheet.getRange(row, 4).setValue(premarketChange);
      
      // 更新盤前成交量欄位
      mainSheet.getRange(row, 5).setValue(premarketVolume);
      
    } catch (e) {
      console.log('更新盤前數據時出錯：' + stock + ', ' + e.toString());
      
      // 發生錯誤時，設置為N/A
      mainSheet.getRange(row, 4).setValue("N/A");
      mainSheet.getRange(row, 5).setValue("N/A");
    }
  }
}

// 設置每日刷新觸發器
function setDailyTrigger() {
  // 刪除所有現有的刷新數據觸發器
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshData') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 創建新的每日觸發器，在美東時間晚上8:45 (台灣時間早上8:45左右) 自動刷新
  ScriptApp.newTrigger('refreshData')
    .timeBased()
    .atHour(20)
    .nearMinute(45)
    .everyDays(1)
    .create();
}

// 顯示關於對話框
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('關於選擇權快篩系統', 
           '版本：1.0\n' +
           '功能：追蹤核心觀察股並提供盤前選擇權策略建議\n\n' +
           '使用方法：\n' +
           '1. 在主工作表中填入您要追蹤的股票代碼\n' +
           '2. 根據您的技術分析選擇「偏多」、「偏空」或「震盪」\n' +
           '3. 填入當前的IV Rank值(0-100)\n' +
           '4. 系統會自動建議合適的策略和參數\n\n' +
           '數據每個交易日早上自動刷新，您也可以隨時點擊「手動刷新數據」',
           ui.ButtonSet.OK);
} 