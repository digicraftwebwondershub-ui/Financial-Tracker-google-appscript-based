// Function to serve the HTML webpage
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Function to get data for the Dashboard
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transactionSheet = ss.getSheetByName("Transactions");
  const creditCardSheet = ss.getSheetByName("Credit Cards");
  const goalsSheet = ss.getSheetByName("Goals");

  const today = new Date();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  const prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;
  const prevMonthYear = currentMonth === 0 ? currentYear - 1 : currentYear;

  // Monthly Summary
  let totalIncome = 0;
  let totalExpenses = 0;
  let totalSavings = 0;
  let prevTotalExpenses = 0;
  let prevTotalSavings = 0;
  const spendingCategories = {};
  const incomeExpenseTrend = { labels: [], income: [], expenses: [] };

  const transactionData = transactionSheet.getDataRange().getValues();
  for (let i = 1; i < transactionData.length; i++) {
    const rowDate = new Date(transactionData[i][0]);
    const month = rowDate.getMonth();
    const year = rowDate.getFullYear();

    if (year === currentYear && month === currentMonth) {
      if (transactionData[i][1] === "Income") {
        totalIncome += transactionData[i][3];
      } else if (transactionData[i][1] === "Expense") {
        totalExpenses += transactionData[i][3];
        const category = transactionData[i][2];
        spendingCategories[category] = (spendingCategories[category] || 0) + transactionData[i][3];
      } else if (transactionData[i][1] === "Savings") {
        totalSavings += transactionData[i][3];
      }
      const day = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MM/dd");
      if (!incomeExpenseTrend.labels.includes(day)) {
        incomeExpenseTrend.labels.push(day);
        incomeExpenseTrend.income.push(0);
        incomeExpenseTrend.expenses.push(0);
      }
      const index = incomeExpenseTrend.labels.indexOf(day);
      if (transactionData[i][1] === "Income") {
        incomeExpenseTrend.income[index] += transactionData[i][3];
      } else if (transactionData[i][1] === "Expense") {
        incomeExpenseTrend.expenses[index] += transactionData[i][3];
      }
    } else if (year === prevMonthYear && month === prevMonth) {
      if (transactionData[i][1] === "Expense") {
        prevTotalExpenses += transactionData[i][3];
      } else if (transactionData[i][1] === "Savings") {
        prevTotalSavings += transactionData[i][3];
      }
    }
  }

  // Credit Card Summary
  let creditUsage = 0;
  let totalLimit = 0;
  let totalBalance = 0;
  const creditCardData = creditCardSheet.getDataRange().getValues();
  for (let i = 1; i < creditCardData.length; i++) {
    totalBalance += creditCardData[i][2];
    totalLimit += creditCardData[i][1];
  }
  creditUsage = (totalBalance / totalLimit) * 100;

  // Goals Summary
  const goalsData = goalsSheet.getDataRange().getValues();
  const goalSummary = [];
  if (goalsData.length > 1) {
    goalsData.slice(1).forEach(goal => {
      const savedAmount = goal[2];
      const targetAmount = goal[1];
      const progress = (savedAmount / targetAmount) * 100;
      const remaining = targetAmount - savedAmount;
      goalSummary.push({
        name: goal[0],
        progress: progress.toFixed(2),
        saved: savedAmount,
        target: targetAmount,
        remaining: remaining,
        dueDate: Utilities.formatDate(new Date(goal[3]), Session.getScriptTimeZone(), "MMM dd, yyyy")
      });
    });
  }

  return {
    netIncome: totalIncome - totalExpenses,
    totalExpenses: totalExpenses,
    savingsRate: (totalSavings / totalIncome) * 100 || 0,
    creditUsage: creditUsage.toFixed(2) || 0,
    totalCreditAvailable: totalLimit - totalBalance,
    creditCardSummary: getCreditCardData(),
    goalsSummary: goalSummary,
    savingsTrend: totalSavings - prevTotalSavings,
    expenseTrend: totalExpenses - prevTotalExpenses,
    spendingCategories: spendingCategories,
    incomeExpenseTrend: incomeExpenseTrend
  };
}

// Function to add a new transaction
function addTransaction(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  const rowData = [
    new Date(formData.date),
    formData.type,
    formData.category,
    parseFloat(formData.amount),
    formData.description,
    formData.paymentMethod,
    formData.accountName
  ];
  sheet.appendRow(rowData);
  return { status: "success", message: "Transaction added successfully!" };
}

// Function to get all transactions
function getTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  const data = sheet.getDataRange().getValues();
  const transactions = [];
  for (let i = 1; i < data.length; i++) {
    transactions.push({
      row: i + 1, // Add row number for editing/deleting
      date: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "MMM dd, yyyy"),
      type: data[i][1],
      category: data[i][2],
      amount: data[i][3],
      description: data[i][4],
      paymentMethod: data[i][5],
      accountName: data[i][6]
    });
  }
  return transactions;
}

// Function to update a transaction
function updateTransaction(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  const row = parseInt(rowData.row);
  sheet.getRange(row, 1, 1, 7).setValues([[
    new Date(rowData.date),
    rowData.type,
    rowData.category,
    parseFloat(rowData.amount),
    rowData.description,
    rowData.paymentMethod,
    rowData.accountName
  ]]);
  return { status: "success", message: "Transaction updated successfully!" };
}

// Function to delete a transaction
function deleteTransaction(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  sheet.deleteRow(parseInt(row));
  return { status: "success", message: "Transaction deleted successfully!" };
}

// Function to add a credit card
function addCreditCard(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Credit Cards");
  const rowData = [
    formData.cardName,
    parseFloat(formData.limit),
    parseFloat(formData.balance),
    parseFloat(formData.apr),
    new Date(formData.dueDate),
    formData.lastPayment ? parseFloat(formData.lastPayment) : '',
    formData.lastPaymentDate ? new Date(formData.lastPaymentDate) : ''
  ];
  sheet.appendRow(rowData);
  return { status: "success", message: "Credit Card added successfully!" };
}

// Function to handle edit credit card data
function editCreditCard(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Credit Cards");
  const row = parseInt(formData.row);
  if (row > 0) {
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setValues([[
      formData.cardName,
      parseFloat(formData.limit),
      parseFloat(formData.balance),
      parseFloat(formData.apr),
      new Date(formData.dueDate),
      parseFloat(formData.lastPayment),
      new Date(formData.lastPaymentDate)
    ]]);
    return { status: "success", message: "Credit card updated successfully!" };
  }
  return { status: "error", message: "Invalid row number." };
}

// Function to handle delete credit card
function deleteCreditCard(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Credit Cards");
  if (row > 0) {
    sheet.deleteRow(row);
    return { status: "success", message: "Credit card deleted successfully!" };
  }
  return { status: "error", message: "Invalid row number." };
}

// Updated getCreditCardData to include status
function getCreditCardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Credit Cards");
  const data = sheet.getDataRange().getValues();
  const cards = [];
  const today = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const balance = data[i][2];
    const limit = data[i][1];
    const available = limit - balance;
    const utilization = (balance / limit) * 100;
    const dueDate = new Date(data[i][4]);
    const daysUntilDue = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));
    
    let status = "";
    if (daysUntilDue < 0 && balance > 0) {
      status = "Overdue";
    } else if (daysUntilDue >= 0 && daysUntilDue <= 7 && balance > 0) {
      status = "Upcoming";
    } else {
      status = "Good";
    }
    
    // Check if Last Payment and Last Payment Date exist before accessing
    const lastPayment = data[i][5] || 0;
    const lastPaymentDate = data[i][6] ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), "MMM dd, yyyy") : 'N/A';

    cards.push({
      row: i + 1,
      name: data[i][0],
      limit: limit,
      balance: balance,
      available: available,
      apr: data[i][3],
      dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "MMM dd, yyyy"),
      lastPayment: lastPayment,
      lastPaymentDate: lastPaymentDate,
      utilization: utilization.toFixed(2),
      status: status
    });
  }
  return cards;
}

// Function to add a savings goal
function addSavingsGoal(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Goals");
  const rowData = [
    formData.goalName,
    parseFloat(formData.targetAmount),
    parseFloat(formData.savedAmount),
    new Date(formData.targetDate)
  ];
  sheet.appendRow(rowData);
  return { status: "success", message: "Savings Goal added successfully!" };
}

// Function to get all savings goals
function getGoalsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Goals");
  const data = sheet.getDataRange().getValues();
  const goals = [];
  for (let i = 1; i < data.length; i++) {
    const targetAmount = data[i][1];
    const savedAmount = data[i][2];
    const targetDate = new Date(data[i][3]);
    const today = new Date();
    const remainingDays = Math.ceil((targetDate - today) / (1000 * 60 * 60 * 24));
    const remainingAmount = targetAmount - savedAmount;
    const monthlySavingsNeeded = remainingDays > 0 ? (remainingAmount / (remainingDays / 30.44)).toFixed(2) : 0;
    const progressPercentage = (savedAmount / targetAmount) * 100;

    goals.push({
      name: data[i][0],
      targetAmount: targetAmount,
      savedAmount: savedAmount,
      remainingAmount: remainingAmount,
      targetDate: Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "MMM dd, yyyy"),
      remainingDays: remainingDays,
      monthlySavingsNeeded: monthlySavingsNeeded,
      progressPercentage: progressPercentage.toFixed(2),
      status: (remainingDays <= 0 && remainingAmount > 0) ? "Overdue" : ""
    });
  }
  return goals;
}

// Function to add a reminder
function addReminder(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reminders");
  const rowData = [
    formData.description,
    new Date(formData.dueDate),
    parseFloat(formData.amount),
    formData.recurring
  ];
  sheet.appendRow(rowData);
  return { status: "success", message: "Reminder added successfully!" };
}

// Function to get all reminders
function getRemindersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reminders");
  const data = sheet.getDataRange().getValues();
  const reminders = [];
  for (let i = 1; i < data.length; i++) {
    const dueDate = new Date(data[i][1]);
    const today = new Date();
    const daysOverdue = Math.ceil((today - dueDate) / (1000 * 60 * 60 * 24));
    const isOverdue = daysOverdue > 0;

    reminders.push({
      description: data[i][0],
      dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "MMM dd, yyyy"),
      amount: data[i][2],
      recurring: data[i][3],
      daysOverdue: daysOverdue,
      isOverdue: isOverdue
    });
  }
  return reminders;
}
