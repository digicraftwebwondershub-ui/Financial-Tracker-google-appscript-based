// Function to serve the HTML webpage
function doGet() {
  createDailyTrigger();
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Function to get data for the Dashboard
function generateUniqueId() {
  return Utilities.getUuid();
}

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
    const rowDate = new Date(transactionData[i][1]);
    const month = rowDate.getMonth();
    const year = rowDate.getFullYear();

    if (year === currentYear && month === currentMonth) {
      if (transactionData[i][2] === "Income") {
        totalIncome += transactionData[i][4];
      } else if (transactionData[i][2] === "Expense") {
        totalExpenses += transactionData[i][4];
        const category = transactionData[i][3];
        spendingCategories[category] = (spendingCategories[category] || 0) + transactionData[i][4];
      } else if (transactionData[i][2] === "Savings") {
        totalSavings += transactionData[i][4];
      }
      const day = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MM/dd");
      if (!incomeExpenseTrend.labels.includes(day)) {
        incomeExpenseTrend.labels.push(day);
        incomeExpenseTrend.income.push(0);
        incomeExpenseTrend.expenses.push(0);
      }
      const index = incomeExpenseTrend.labels.indexOf(day);
      if (transactionData[i][2] === "Income") {
        incomeExpenseTrend.income[index] += transactionData[i][4];
      } else if (transactionData[i][2] === "Expense") {
        incomeExpenseTrend.expenses[index] += transactionData[i][4];
      }
    } else if (year === prevMonthYear && month === prevMonth) {
      if (transactionData[i][2] === "Expense") {
        prevTotalExpenses += transactionData[i][4];
      } else if (transactionData[i][2] === "Savings") {
        prevTotalSavings += transactionData[i][4];
      }
    }
  }

  // Credit Card Summary
  let creditUsage = 0;
  let totalLimit = 0;
  let totalBalance = 0;
  const creditCardData = creditCardSheet.getDataRange().getValues();
  for (let i = 1; i < creditCardData.length; i++) {
    totalBalance += creditCardData[i][4];
    totalLimit += creditCardData[i][3];
  }
  creditUsage = (totalBalance / totalLimit) * 100;

  // Goals Summary
  const goalsData = goalsSheet.getDataRange().getValues();
  const goalSummary = [];
  if (goalsData.length > 1) {
    goalsData.slice(1).forEach(goal => {
      const savedAmount = goal[4];
      const targetAmount = goal[3];
      const progress = (savedAmount / targetAmount) * 100;
      const remaining = targetAmount - savedAmount;
      goalSummary.push({
        name: goal[1],
        progress: progress.toFixed(2),
        saved: savedAmount,
        target: targetAmount,
        remaining: remaining,
        dueDate: Utilities.formatDate(new Date(goal[5]), Session.getScriptTimeZone(), "MMM dd, yyyy")
      });
    });
  }

  const savingsImprovement = ((totalSavings - prevTotalSavings) / prevTotalSavings) * 100;
  let motivationalMessage = "Keep going! Every peso saved is a step toward freedom.";
  if (savingsImprovement > 10) {
    motivationalMessage = `You saved ${savingsImprovement.toFixed(0)}% more this month than last — amazing progress!`;
  } else if (totalExpenses > totalIncome) {
    motivationalMessage = "It’s okay if you overspent — awareness is the first step to control.";
  }

  const tips = [
    "💡 Tip of the week: Automate your savings to build consistency.",
    "💡 Tip of the week: Review your subscriptions. Do you use them all?",
    "💡 Tip of the week: Try a no-spend weekend challenge!"
  ];
  const weeklyTip = tips[Math.floor(Math.random() * tips.length)];

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
    incomeExpenseTrend: incomeExpenseTrend,
    welcomeMessage: `Good morning, Michelle! Here’s how your finances look today ☀️`,
    motivationalMessage: motivationalMessage,
    weeklyTip: weeklyTip
  };
}

// V2 Functions with Transaction Linking Logic

function addTransaction(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transactionSheet = ss.getSheetByName("Transactions");
  const amount = parseFloat(formData.amount);

  // 1. Add to Transactions Sheet
  const newRowData = [
    generateUniqueId(),
    new Date(formData.date),
    formData.type,
    formData.category,
    amount,
    formData.description,
    formData.paymentMethod,
    formData.accountName,
    formData.cardId || null,
    formData.goalId || null
  ];
  transactionSheet.appendRow(newRowData);

  // 2. Update Linked Sheets
  if (formData.type === 'Savings' && formData.goalId) {
    updateGoalProgress(formData.goalId, amount);
  }

  return { status: 'success', message: 'Transaction added and linked successfully!' };
}

function updateTransaction(newFormData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  const row = parseInt(newFormData.row);
  
  // 1. Get Old Transaction Data to Revert Changes
  const oldRowData = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const oldTransaction = {
    type: oldRowData[2],
    category: oldRowData[3],
    amount: parseFloat(oldRowData[4]),
    paymentMethod: oldRowData[6],
    cardId: oldRowData[8],
    goalId: oldRowData[9]
  };

  // 2. Revert the old transaction's impact
  if (oldTransaction.type === 'Savings' && oldTransaction.goalId) {
    updateGoalProgress(oldTransaction.goalId, -oldTransaction.amount);
  }

  // 3. Apply the new transaction's impact
  const newAmount = parseFloat(newFormData.amount);
  if (newFormData.type === 'Savings' && newFormData.goalId) {
    updateGoalProgress(newFormData.goalId, newAmount);
  }
  
  // 4. Update the transaction row in the sheet
  sheet.getRange(row, 1, 1, 10).setValues([[
    newFormData.transactionId,
    new Date(newFormData.date),
    newFormData.type,
    newFormData.category,
    newAmount,
    newFormData.description,
    newFormData.paymentMethod,
    newFormData.accountName,
    newFormData.cardId,
    newFormData.goalId
  ]]);

  return { status: 'success', message: 'Transaction updated successfully!' };
}

function getTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  const data = sheet.getDataRange().getValues();
  const transactions = [];
  for (let i = 1; i < data.length; i++) {
    transactions.push({
      row: i + 1, // Add row number for editing/deleting
      transactionId: data[i][0],
      date: new Date(data[i][1]).toISOString(),
      type: data[i][2],
      category: data[i][3],
      amount: data[i][4],
      description: data[i][5],
      paymentMethod: data[i][6],
      accountName: data[i][7],
      cardId: data[i][8],
      goalId: data[i][9]
    });
  }
  return transactions;
}

function deleteTransaction(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");

  // 1. Get Data to Revert Changes
  const rowData = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const transaction = {
    type: rowData[2],
    category: rowData[3],
    amount: parseFloat(rowData[4]),
    paymentMethod: rowData[6],
    cardId: rowData[8],
    goalId: rowData[9]
  };

  // 2. Revert the transaction's impact
  if (transaction.type === 'Savings' && transaction.goalId) {
    updateGoalProgress(transaction.goalId, -transaction.amount);
  }

  // 3. Delete the row
  sheet.deleteRow(row);
  return { status: 'success', message: 'Transaction deleted successfully!' };
}

// Helper function to update goal progress
function updateGoalProgress(goalId, amount) {
  const goalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Goals");
  const data = goalSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === goalId) {
      const currentSaved = data[i][4];
      const newSaved = currentSaved + amount;
      goalSheet.getRange(i + 1, 5).setValue(newSaved);
      break;
    }
  }
}

// Function to add a credit card
function addCreditCard(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Credit Cards");
  const rowData = [
    generateUniqueId(),
    formData.cardName,
    formData.bankName,
    parseFloat(formData.limit),
    parseFloat(formData.balance),
    formData.lastPayment ? parseFloat(formData.lastPayment) : '',
    parseFloat(formData.apr),
    formData.lastPaymentDate ? new Date(formData.lastPaymentDate) : '',
    new Date(formData.dueDate),
    formData.statementDate ? new Date(formData.statementDate) : ''
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
      formData.cardId,
      formData.cardName,
      formData.bankName,
      parseFloat(formData.limit),
      parseFloat(formData.balance),
      parseFloat(formData.lastPayment),
      parseFloat(formData.apr),
      new Date(formData.lastPaymentDate),
      new Date(formData.dueDate),
      new Date(formData.statementDate)
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
  const cardSheet = ss.getSheetByName("Credit Cards");
  const transactionSheet = ss.getSheetByName("Transactions");

  const cardData = cardSheet.getDataRange().getValues();
  const transactionData = transactionSheet.getDataRange().getValues();
  const cards = [];
  const today = new Date();

  // Pre-calculate transaction totals and find the last payment for each card
  const cardTransactionTotals = {};
  const lastPayments = {};
  for (let i = 1; i < transactionData.length; i++) {
    const cardId = transactionData[i][8];
    if (cardId) {
      if (!cardTransactionTotals[cardId]) {
        cardTransactionTotals[cardId] = 0;
      }
      const type = String(transactionData[i][2]).trim();
      const category = String(transactionData[i][3]).trim();
      const amount = parseFloat(transactionData[i][4]);

      if (category.toLowerCase() === 'credit card payment') {
        cardTransactionTotals[cardId] -= amount;

        const paymentDate = new Date(transactionData[i][1]);
        if (!lastPayments[cardId] || paymentDate > lastPayments[cardId].date) {
          lastPayments[cardId] = {
            amount: amount,
            date: paymentDate
          };
        }
      } else if (type.toLowerCase() === 'expense') {
        cardTransactionTotals[cardId] += amount;
      }
    }
  }
  
  for (let i = 1; i < cardData.length; i++) {
    const cardId = cardData[i][0];
    const balance = cardTransactionTotals[cardId] || 0;

    const limit = cardData[i][3];
    const available = limit - balance;
    const utilization = (balance / limit) * 100;
    const dueDate = new Date(cardData[i][8]);
    const daysUntilDue = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));

    let status = "";
    if (daysUntilDue < 0 && balance > 0) {
      status = "Overdue";
    } else if (daysUntilDue >= 0 && daysUntilDue <= 7 && balance > 0) {
      status = "Upcoming";
    } else {
      status = "Good";
    }

    const lastPaymentInfo = lastPayments[cardId];
    const lastPayment = lastPaymentInfo ? lastPaymentInfo.amount : 0;
    const lastPaymentDate = lastPaymentInfo ? Utilities.formatDate(lastPaymentInfo.date, Session.getScriptTimeZone(), "MMM dd, yyyy") : 'N/A';
    const bankName = cardData[i][2] || '';
    const statementDate = cardData[i][9] ? Utilities.formatDate(new Date(cardData[i][9]), Session.getScriptTimeZone(), "MMM dd") : 'N/A';

    let insight = "";
    if (daysUntilDue > 0 && daysUntilDue <= 5) {
      insight = `✅ Your next bill is due in ${daysUntilDue} days — you’ve got this!`;
    } else if (utilization > 70) {
      const amountToPay = (balance - (limit * 0.3)).toFixed(2);
      insight = `⚠️ You’re close to your limit. Try paying ₱${amountToPay} to bring usage under 30%.`;
    } else if (lastPayment > 0) {
      insight = "👏 No missed payments this month. Keep your streak!";
    }

    cards.push({
      row: i + 1,
      cardId: cardId,
      name: cardData[i][1],
      limit: limit,
      balance: balance,
      available: available,
      apr: cardData[i][6],
      dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "MMM dd, yyyy"),
      lastPayment: lastPayment,
      lastPaymentDate: lastPaymentDate,
      utilization: utilization.toFixed(2),
      status: status,
      bankName: bankName,
      statementDate: statementDate,
      insight: insight
    });
  }
  return cards;
}

// Function to add a savings goal
function addSavingsGoal(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Goals");
  const rowData = [
    generateUniqueId(),
    formData.goalName,
    formData.category,
    parseFloat(formData.targetAmount),
    parseFloat(formData.savedAmount),
    new Date(formData.targetDate),
    '', // Placeholder for monthly savings, can be calculated or entered manually
    formData.priority
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
    const targetAmount = data[i][3];
    const savedAmount = data[i][4];
    const targetDate = new Date(data[i][5]);
    const today = new Date();
    const remainingDays = Math.ceil((targetDate - today) / (1000 * 60 * 60 * 24));
    const remainingAmount = targetAmount - savedAmount;
    const monthlySavingsNeeded = remainingDays > 0 ? (remainingAmount / (remainingDays / 30.44)).toFixed(2) : 0;
    const progressPercentage = (savedAmount / targetAmount) * 100;
    
    let insight = "";
    if (progressPercentage >= 100) {
      insight = "✨ Goal achieved! Treat yourself (mindfully).";
    } else if (progressPercentage >= 70) {
      insight = `🎯 You’re ${progressPercentage.toFixed(0)}% to your ${data[i][1]} fund — only ₱${remainingAmount.toLocaleString()} to go!`;
    } else if (monthlySavingsNeeded > 0) {
      insight = `💪 Stay consistent! Saving ₱${(monthlySavingsNeeded / 4).toLocaleString()} more this week gets you back on track.`;
    }

    goals.push({
      goalId: data[i][0],
      name: data[i][1],
      targetAmount: targetAmount,
      savedAmount: savedAmount,
      remainingAmount: remainingAmount,
      targetDate: Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "MMM dd, yyyy"),
      remainingDays: remainingDays,
      monthlySavingsNeeded: monthlySavingsNeeded,
      progressPercentage: progressPercentage.toFixed(2),
      status: (remainingDays <= 0 && remainingAmount > 0) ? "Overdue" : "",
      category: data[i][2],
      priority: data[i][7],
      insight: insight
    });
  }
  return goals;
}

function getCardsForDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Credit Cards");
  const data = sheet.getDataRange().getValues();
  const cards = [];
  for (let i = 1; i < data.length; i++) {
    cards.push({
      id: data[i][0],
      name: data[i][1]
    });
  }
  return cards;
}

function getGoalsForDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Goals");
  const data = sheet.getDataRange().getValues();
  const goals = [];
  for (let i = 1; i < data.length; i++) {
    goals.push({
      id: data[i][0],
      name: data[i][1]
    });
  }
  return goals;
}

// Function to add a reminder
function addReminder(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reminders");
  const rowData = [
    generateUniqueId(),
    formData.description,
    formData.category,
    new Date(formData.dueDate),
    parseFloat(formData.amount),
    'Pending', // Default status
    formData.recurring,
    '', // Placeholder for days left,
    formData.paymentChannel,
    formData.autoNotify
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
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    const dueDate = new Date(data[i][3]);
    dueDate.setHours(0, 0, 0, 0);
    const daysLeft = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));
    const status = data[i][5];

    let urgency = 'safe';
    if (daysLeft < 3 || status === 'Overdue') {
      urgency = 'urgent';
    } else if (daysLeft <= 7) {
      urgency = 'approaching';
    }

    let feedback = "";
    if (status === 'Paid') {
      feedback = "✅ You paid your " + data[i][1] + " on time — nice discipline!";
    } else if (daysLeft < 0) {
      feedback = "⏰ Overdue! Don’t stress — mark it done once paid.";
    } else if (daysLeft <= 2) {
      feedback = "📅 " + data[i][1] + " bill due in " + daysLeft + " days — schedule a reminder payment.";
    }

    reminders.push({
      reminderId: data[i][0],
      description: data[i][1],
      dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "MMM dd, yyyy"),
      amount: data[i][4],
      status: status,
      recurring: data[i][6],
      daysLeft: daysLeft,
      urgency: urgency,
      feedback: feedback,
      category: data[i][2],
      autoNotify: data[i][9],
      paymentChannel: data[i][8]
    });
  }
  return reminders;
}

function updateDaysLeft() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reminders");
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    const dueDate = new Date(data[i][3]);
    dueDate.setHours(0, 0, 0, 0);
    const daysLeft = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));
    sheet.getRange(i + 1, 8).setValue(daysLeft);
  }
}

function createDailyTrigger() {
  // Deletes all existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }

  // Creates a new trigger
  ScriptApp.newTrigger('updateDaysLeft')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
}
