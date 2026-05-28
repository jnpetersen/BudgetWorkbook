const CURRENT_VERSION = "1.0.0";
const UPDATE_INFO_URL = "https://raw.githubusercontent.com/jnpetersen/BudgetWorkbook/main/version.json"

function checkForUpdates() {
  const ui = SpreadsheetApp.getUi();

  const response = UrlFetchApp.fetch(UPDATE_INFO_URL, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    ui.alert("Could not check for updates.");
    return;
  }

  const updateInfo = JSON.parse(response.getContentText());
  const latestVersion = updateInfo.version;

  if (isNewerVersion(latestVersion, CURRENT_VERSION)) {
    ui.alert("Update Available",
    `Current version: ${CURRENT_VERSION}\nLatest version: ${latestVersion}\n\nDownload update here:\n${updateInfo.downloadUrl}`,
    ui.ButtonSet.OK
    );
  } else {
    ui.alert("You are already on the latest version.");
  }
}

function isNewerVersion(latest, current) {
  const latestParts = latest.split(".").map(Number);
  const currentParts = current.split(".").map(Number);

  for (let i = 0; i < Math.max(latestParts.length, currentParts.length); i++) {
    const latestNum = latestParts[i] || 0;
    const currentNum = currentParts[i] || 0;

    if (latestNum > currentNum) return true;
    if (latestNum < currentNum) return false;
  }

  return false;
}


function rebuildBudgetPlanner() {
  /****************************************************
   * BASIC SETUP
   ****************************************************/
  const DEBUG_TOASTS = false;

  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName("Settings");
  const planner = ss.getSheetByName("Budget Planner");

  const yearCell = "F1";
  const yearRef = "$F$1";
  const startRow = 8;

  const summaryStartRow = 3;
  const totalIncomeRow = 4;
  const totalExpensesRow = 5;
  const netIncomeRow = 6;

  ss.toast("Rebuilding Budget Planner...", "Please Wait", 10);
  planner.getRange("A2").setValue("Rebuilding Budget Planner...");
  SpreadsheetApp.flush();


  /****************************************************
   * MONTH CONFIGURATION
   ****************************************************/
  const months = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  const firstMonthCol = 2; // Column B
  const columnsPerMonth = 3;

  const totalBudgetCol = firstMonthCol + months.length * columnsPerMonth;
  const totalActualCol = totalBudgetCol + 1;
  const totalOffsetCol = totalBudgetCol + 2;

  const totalPlannerCols = 1 + months.length * columnsPerMonth + 3;


  /****************************************************
   * READ SECTION LIST FROM SETTINGS
   *
   * C = Section Name
   * D = Section Sort Order
   * E = Active
   ****************************************************/
  const lastSettingsRow = settings.getLastRow();

  if (lastSettingsRow < 2) {
    ss.toast("No Settings rows found.", "Stopped", 5);
    planner.getRange("A2").setValue("No Settings rows found.");
    return;
  }

  const sectionData = settings
    .getRange(2, 3, lastSettingsRow - 1, 3)
    .getValues();

  const sectionOrder = sectionData
    .filter(row => row[0] && row[2] === true)
    .sort((a, b) => Number(a[1] || 999) - Number(b[1] || 999))
    .map(row => row[0]);


  /****************************************************
   * READ ITEM LIST FROM SETTINGS
   *
   * G = Item Name
   * H = Section Name
   * I = Active
   * J = Item Sort Order
   ****************************************************/
  const itemData = settings
    .getRange(2, 7, lastSettingsRow - 1, 4)
    .getValues();

  const activeItems = itemData
    .filter(row => row[0] && row[1] && row[2] === true)
    .sort((a, b) => Number(a[3] || 999) - Number(b[3] || 999));


  /****************************************************
   * GROUP ITEMS BY SECTION
   ****************************************************/
  const itemsBySection = {};

  activeItems.forEach(row => {
    const itemName = row[0];
    const sectionName = row[1];

    if (!itemsBySection[sectionName]) {
      itemsBySection[sectionName] = [];
    }

    itemsBySection[sectionName].push(itemName);
  });


  /****************************************************
   * SAVE EXISTING BUDGET VALUES BEFORE CLEARING
   *
   * Key format:
   * Section Name|Item Name|Month Number
   ****************************************************/
  const savedBudgets = {};
  const sectionNames = new Set(sectionOrder);
  let currentSection = "";

  const lastPlannerRow = planner.getLastRow();

  if (lastPlannerRow >= startRow) {
    const existingValues = planner
      .getRange(startRow, 1, lastPlannerRow - startRow + 1, totalPlannerCols)
      .getValues();

    existingValues.forEach(rowValues => {
      const nameValue = rowValues[0];

      if (!nameValue) return;

      if (sectionNames.has(nameValue)) {
        currentSection = nameValue;
        return;
      }

      if (String(nameValue).startsWith("Total ")) return;
      if (!currentSection) return;

      months.forEach((monthName, monthIndex) => {
        const monthNumber = monthIndex + 1;
        const budgetColIndex = firstMonthCol - 1 + monthIndex * columnsPerMonth;
        const budgetValue = rowValues[budgetColIndex];

        if (budgetValue !== "" && budgetValue !== null) {
          const key = `${currentSection}|${nameValue}|${monthNumber}`;
          savedBudgets[key] = budgetValue;
        }
      });
    });
  }


  /****************************************************
   * CLEAR EXISTING BUDGET AREA
   ****************************************************/
  planner
    .getRange(startRow, 1, planner.getMaxRows() - startRow + 1, totalPlannerCols)
    .breakApart()
    .clear();


  /****************************************************
   * REMOVE OLD CONDITIONAL FORMAT RULES FOR GENERATED AREA
   ****************************************************/
  const existingRules = planner.getConditionalFormatRules();

  const keptRules = existingRules.filter(rule => {
    return !rule.getRanges().some(range => range.getRow() >= startRow);
  });

  const newConditionalRules = [...keptRules];


  /****************************************************
   * SECTION COLORS / STYLE SETTINGS
   ****************************************************/
  const sectionColors = {
    "INCOME": "#7CB342",
    "SAVINGS EXPENSE": "#78909C",
    "HOME EXPENSES": "#78909C",
    "AUTO EXPENSES": "#78909C",
    "FOOD EXPENSES": "#78909C",
    "PERSONAL EXPENSES": "#78909C"
  };

  const defaultSectionColor = "#78909C";
  const alternatingColor = "#F1F1F1";
  const currentMonthHighlight = "#FFF2CC";

  const positiveOffsetColor = "#b7e1cd";
  const negativeOffsetColor = "#f4c7c3";


  /****************************************************
   * BUILD BUDGET PLANNER
   ****************************************************/
  let currentRow = startRow;
  const sectionTotalRows = {};

  sectionOrder.forEach(sectionName => {
    const sectionStartTime = new Date();
    const items = itemsBySection[sectionName] || [];

    if (items.length === 0) return;

    if (DEBUG_TOASTS) {
      ss.toast(`Building ${sectionName}...`, "Debug", 5);
      SpreadsheetApp.flush();
    }

    const itemStartRow = currentRow + 1;
    const itemEndRow = currentRow + items.length;


    /****************************************************
     * CREATE SECTION HEADER ROW
     ****************************************************/
    planner.getRange(currentRow, 1)
      .setValue(sectionName)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    planner.getRange(currentRow, 1, 1, totalPlannerCols)
      .setBackground(sectionColors[sectionName] || defaultSectionColor)
      .setFontColor("#000000")
      .setFontWeight("bold");


    /****************************************************
     * CREATE MONTH TITLES IN SECTION BAR
     * Budget + Actual are merged with month name.
     * Offset remains separate.
     ****************************************************/
    months.forEach((monthName, index) => {
      const baseCol = firstMonthCol + index * columnsPerMonth;

      planner.getRange(currentRow, baseCol, 1, 2).merge();

      planner.getRange(currentRow, baseCol)
        .setValue(monthName)
        .setHorizontalAlignment("center")
        .setFontWeight("bold");

      planner.getRange(currentRow, baseCol + 2)
        .setValue("Offset")
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
    });


    /****************************************************
     * CREATE YEARLY TOTAL TITLES
     ****************************************************/
    planner.getRange(currentRow, totalBudgetCol).setValue("Budget");
    planner.getRange(currentRow, totalActualCol).setValue("Actual");
    planner.getRange(currentRow, totalOffsetCol).setValue("Offset");

    planner.getRange(currentRow, totalBudgetCol, 1, 3)
      .setHorizontalAlignment("center")
      .setFontWeight("bold");

    currentRow++;


    /****************************************************
     * CREATE ITEM ROWS
     ****************************************************/
    items.forEach((itemName, itemIndex) => {
      planner.getRange(currentRow, 1).setValue(itemName);

      /****************************************************
       * ALTERNATING ITEM ROW COLORS
       ****************************************************/
      if (itemIndex % 2 === 1) {
        planner.getRange(currentRow, 1, 1, totalPlannerCols)
          .setBackground(alternatingColor);
      }

      const budgetCells = [];
      const actualCells = [];
      const offsetCells = [];

      months.forEach((monthName, index) => {
        const monthNumber = index + 1;
        const baseCol = firstMonthCol + index * columnsPerMonth;

        const budgetCell = planner.getRange(currentRow, baseCol);
        const actualCell = planner.getRange(currentRow, baseCol + 1);
        const offsetCell = planner.getRange(currentRow, baseCol + 2);

        budgetCells.push(budgetCell.getA1Notation());
        actualCells.push(actualCell.getA1Notation());
        offsetCells.push(offsetCell.getA1Notation());


        /****************************************************
         * RESTORE SAVED BUDGET VALUE
         ****************************************************/
        const savedKey = `${sectionName}|${itemName}|${monthNumber}`;

        if (savedBudgets[savedKey] !== undefined) {
          budgetCell.setValue(savedBudgets[savedKey]);
        }


        /****************************************************
         * ACTUAL FORMULA
         *
         * Transactions tab expected columns:
         * A = Date
         * E = Category
         * G = Withdrawal / Payment
         * H = Deposit / Credit
         *
         * Actual remains blank when no transactions exist.
         ****************************************************/
        const sumFormula = sectionName === "INCOME"
          ? `SUMIFS(Transactions!$H:$H,Transactions!$E:$E,$A${currentRow},Transactions!$A:$A,">="&DATE(${yearRef},${monthNumber},1),Transactions!$A:$A,"<"&EDATE(DATE(${yearRef},${monthNumber},1),1))`
          : `SUMIFS(Transactions!$G:$G,Transactions!$E:$E,$A${currentRow},Transactions!$A:$A,">="&DATE(${yearRef},${monthNumber},1),Transactions!$A:$A,"<"&EDATE(DATE(${yearRef},${monthNumber},1),1))`;

        actualCell.setFormula(`=IF(${sumFormula}=0,"",${sumFormula})`);


        /****************************************************
         * OFFSET FORMULA
         *
         * Offset always shows.
         * Blank Actual is treated as zero by Google Sheets.
         ****************************************************/
        offsetCell.setFormula(
          `=${budgetCell.getA1Notation()}-${actualCell.getA1Notation()}`
        );

        planner.getRange(currentRow, baseCol, 1, 3)
          .setNumberFormat("$#,##0.00");
      });


      /****************************************************
       * CREATE ITEM YEARLY TOTAL FORMULAS
       ****************************************************/
      planner.getRange(currentRow, totalBudgetCol)
        .setFormula(`=IF(SUM(${budgetCells.join(",")})=0,"",SUM(${budgetCells.join(",")}))`);

      planner.getRange(currentRow, totalActualCol)
        .setFormula(`=IF(SUM(${actualCells.join(",")})=0,"",SUM(${actualCells.join(",")}))`);

      planner.getRange(currentRow, totalOffsetCol)
        .setFormula(`=SUM(${offsetCells.join(",")})`);

      planner.getRange(currentRow, totalBudgetCol, 1, 3)
        .setNumberFormat("$#,##0.00");

      currentRow++;
    });


    /****************************************************
     * CREATE TOTAL ROW FOR SECTION
     ****************************************************/
    planner.getRange(currentRow, 1)
      .setValue("Total " + sectionName)
      .setFontWeight("bold")
      .setHorizontalAlignment("right");

    months.forEach((monthName, index) => {
      const baseCol = firstMonthCol + index * columnsPerMonth;

      planner.getRange(currentRow, baseCol)
        .setFormula(`=SUM(${planner.getRange(itemStartRow, baseCol).getA1Notation()}:${planner.getRange(itemEndRow, baseCol).getA1Notation()})`);

      planner.getRange(currentRow, baseCol + 1)
        .setFormula(`=SUM(${planner.getRange(itemStartRow, baseCol + 1).getA1Notation()}:${planner.getRange(itemEndRow, baseCol + 1).getA1Notation()})`);

      planner.getRange(currentRow, baseCol + 2)
        .setFormula(`=SUM(${planner.getRange(itemStartRow, baseCol + 2).getA1Notation()}:${planner.getRange(itemEndRow, baseCol + 2).getA1Notation()})`);

      planner.getRange(currentRow, baseCol, 1, 3)
        .setNumberFormat("$#,##0.00")
        .setFontWeight("bold");
    });


    /****************************************************
     * CREATE SECTION YEARLY TOTAL FORMULAS
     ****************************************************/
    planner.getRange(currentRow, totalBudgetCol)
      .setFormula(`=SUM(${planner.getRange(itemStartRow, totalBudgetCol).getA1Notation()}:${planner.getRange(itemEndRow, totalBudgetCol).getA1Notation()})`);

    planner.getRange(currentRow, totalActualCol)
      .setFormula(`=SUM(${planner.getRange(itemStartRow, totalActualCol).getA1Notation()}:${planner.getRange(itemEndRow, totalActualCol).getA1Notation()})`);

    planner.getRange(currentRow, totalOffsetCol)
      .setFormula(`=SUM(${planner.getRange(itemStartRow, totalOffsetCol).getA1Notation()}:${planner.getRange(itemEndRow, totalOffsetCol).getA1Notation()})`);

    planner.getRange(currentRow, totalBudgetCol, 1, 3)
      .setNumberFormat("$#,##0.00")
      .setFontWeight("bold");

    sectionTotalRows[sectionName] = currentRow;

    currentRow++;


    /****************************************************
     * BLANK ROW BETWEEN SECTIONS
     ****************************************************/
    currentRow++;


    /****************************************************
     * MERGED CURRENT-MONTH MARKER ROW
     *
     * One merged 3-column cell per month.
     * Conditional formatting highlights the current month.
     ****************************************************/
    months.forEach((monthName, index) => {
      const baseCol = firstMonthCol + index * columnsPerMonth;

      planner.getRange(currentRow, baseCol, 1, columnsPerMonth).merge();

      const markerRange = planner.getRange(currentRow, baseCol, 1, columnsPerMonth);

      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=MONTH(TODAY())=${index + 1}`)
        .setBackground(currentMonthHighlight)
        .setRanges([markerRange])
        .build();

      newConditionalRules.push(rule);
    });

    currentRow++;

    const sectionEndTime = new Date();
    const elapsedSeconds = ((sectionEndTime - sectionStartTime) / 1000).toFixed(2);

    if (DEBUG_TOASTS) {
      ss.toast(`${sectionName} completed in ${elapsedSeconds}s`, "Debug", 5);
      SpreadsheetApp.flush();
    }
  });


  /****************************************************
   * CONDITIONAL FORMAT OFFSET VALUES
   * Green = 0 or above
   * Red = below 0
   ****************************************************/
  months.forEach((monthName, index) => {
    const offsetCol = firstMonthCol + index * columnsPerMonth + 2;

    const offsetRange = planner.getRange(
      startRow,
      offsetCol,
      currentRow - startRow,
      1
    );

    newConditionalRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0)
        .setBackground(positiveOffsetColor)
        .setRanges([offsetRange])
        .build()
    );

    newConditionalRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground(negativeOffsetColor)
        .setRanges([offsetRange])
        .build()
    );
  });


  /****************************************************
   * APPLY CONDITIONAL FORMAT RULES
   ****************************************************/
  planner.setConditionalFormatRules(newConditionalRules);


  /****************************************************
   * UPDATE SUMMARY SECTION
   *
   * Row 4 = Total Income
   * Row 5 = Total Expenses
   * Row 6 = Net Income - Expenses
   *
   * Income includes prior month net carry-forward
   * for months in the current year up to current month.
   ****************************************************/
  months.forEach((monthName, index) => {
    const baseCol = firstMonthCol + index * columnsPerMonth;

    const incomeRow = sectionTotalRows["INCOME"];
    const expenseRows = Object.keys(sectionTotalRows)
      .filter(sectionName => sectionName !== "INCOME")
      .map(sectionName => sectionTotalRows[sectionName]);

    const incomeBudgetCell = incomeRow
      ? planner.getRange(incomeRow, baseCol).getA1Notation()
      : "0";

    const incomeActualCell = incomeRow
      ? planner.getRange(incomeRow, baseCol + 1).getA1Notation()
      : "0";

    const expenseBudgetCells = expenseRows.map(row =>
      planner.getRange(row, baseCol).getA1Notation()
    );

    const expenseActualCells = expenseRows.map(row =>
      planner.getRange(row, baseCol + 1).getA1Notation()
    );

    const expenseBudgetFormula = expenseBudgetCells.length
      ? `SUM(${expenseBudgetCells.join(",")})`
      : "0";

    const expenseActualFormula = expenseActualCells.length
      ? `SUM(${expenseActualCells.join(",")})`
      : "0";


    /****************************************************
     * CARRY FORWARD PREVIOUS MONTH NET INTO INCOME BUDGET
     ****************************************************/
    let carryForwardFormula = "";

    if (index > 0) {
      const previousNetBudgetCell = planner
        .getRange(netIncomeRow, baseCol - columnsPerMonth)
        .getA1Notation();

      carryForwardFormula =
        `+IF(YEAR(DATE(${yearRef},${index + 1},1))=YEAR(TODAY()),IF(MONTH(DATE(${yearRef},${index + 1},1))<=MONTH(TODAY()),${previousNetBudgetCell},0),0)`;
    }

    planner.getRange(totalIncomeRow, baseCol)
      .setFormula(`=${incomeBudgetCell}${carryForwardFormula}`);

    planner.getRange(totalIncomeRow, baseCol + 1)
      .setFormula(`=${incomeActualCell}`);

    planner.getRange(totalIncomeRow, baseCol + 2)
      .setFormula(`=${planner.getRange(totalIncomeRow, baseCol).getA1Notation()}-${planner.getRange(totalIncomeRow, baseCol + 1).getA1Notation()}`);

    planner.getRange(totalExpensesRow, baseCol)
      .setFormula(`=${expenseBudgetFormula}`);

    planner.getRange(totalExpensesRow, baseCol + 1)
      .setFormula(`=${expenseActualFormula}`);

    planner.getRange(totalExpensesRow, baseCol + 2)
      .setFormula(`=${planner.getRange(totalExpensesRow, baseCol).getA1Notation()}-${planner.getRange(totalExpensesRow, baseCol + 1).getA1Notation()}`);

    planner.getRange(netIncomeRow, baseCol)
      .setFormula(`=${planner.getRange(totalIncomeRow, baseCol).getA1Notation()}-${planner.getRange(totalExpensesRow, baseCol).getA1Notation()}`);

    planner.getRange(netIncomeRow, baseCol + 1)
      .setFormula(`=${planner.getRange(totalIncomeRow, baseCol + 1).getA1Notation()}-${planner.getRange(totalExpensesRow, baseCol + 1).getA1Notation()}`);

    planner.getRange(netIncomeRow, baseCol + 2)
      .setFormula(`=${planner.getRange(netIncomeRow, baseCol).getA1Notation()}-${planner.getRange(netIncomeRow, baseCol + 1).getA1Notation()}`);

    planner.getRange(totalIncomeRow, baseCol, 3, 3)
      .setNumberFormat("$#,##0.00");
  });


  /****************************************************
   * SUMMARY YEARLY TOTALS
   ****************************************************/
  [totalIncomeRow, totalExpensesRow, netIncomeRow].forEach(row => {
    const budgetCells = [];
    const actualCells = [];
    const offsetCells = [];

    months.forEach((monthName, index) => {
      const baseCol = firstMonthCol + index * columnsPerMonth;

      budgetCells.push(planner.getRange(row, baseCol).getA1Notation());
      actualCells.push(planner.getRange(row, baseCol + 1).getA1Notation());
      offsetCells.push(planner.getRange(row, baseCol + 2).getA1Notation());
    });

    planner.getRange(row, totalBudgetCol).setFormula(`=SUM(${budgetCells.join(",")})`);
    planner.getRange(row, totalActualCol).setFormula(`=SUM(${actualCells.join(",")})`);
    planner.getRange(row, totalOffsetCol).setFormula(`=SUM(${offsetCells.join(",")})`);

    planner.getRange(row, totalBudgetCol, 1, 3)
      .setNumberFormat("$#,##0.00")
      .setFontWeight("bold");
  });


  /****************************************************
   * FINAL FORMATTING
   ****************************************************/
  if (currentRow > startRow) {
    planner.getRange(startRow, 1, currentRow - startRow, totalPlannerCols)
      .setFontFamily("Calibri")
      .setFontSize(10)
      .setVerticalAlignment("middle");
  }

  planner.getRange(summaryStartRow, 1, 4, totalPlannerCols)
    .setFontFamily("Calibri")
    .setFontSize(10)
    .setVerticalAlignment("middle");


  /****************************************************
   * COLUMN WIDTHS
   ****************************************************/
  planner.setColumnWidth(1, 220);

  for (let col = 2; col <= totalPlannerCols; col++) {
    planner.setColumnWidth(col, 65);
  }


  /****************************************************
   * FINISH
   ****************************************************/
  planner.getRange("A2").setValue("Last Rebuilt: " + new Date());
  ss.toast("Budget Planner rebuild complete.", "Finished", 5);
}