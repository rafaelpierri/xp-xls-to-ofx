XLSX = require('xlsx');
aguid = require('aguid');
handlebars = require('handlebars');
fs = require('fs');

fs.readFile('./templates/bank-statement.hbs', 'utf8', function (err,data) {
  if (err) {
    return console.log(err);
  }

  Date.prototype.toOFXDate = function(){
    return this.getFullYear() +
      ("0" + (this.getMonth() + 1)).slice(-2) +
      ("0" + this.getDate()).slice(-2) +
      ("0" + this.getHours()).slice(-2) +
      ("0" + this.getMinutes()).slice(-2) +
      ("0" + this.getSeconds()).slice(-2) +
      "[" + this.getTimezoneOffset() / 60 * - 1 + ":GMT]";
  }

  var bankStatementTemplate = handlebars.compile(data);
  
  var workbook = XLSX.readFile('extrato.xls');
  
  Date.prototype.addDays=function(d){return new Date(this.valueOf()+864E5*d);};
  
  var date = new Date(1900, 0, 1, 0, 0, 0, 0);
  
  var bankStatement = {};
  var index = 25;
  var transactions = [];
  
  bankStatement.dateTaken = (new Date(workbook.Sheets.Sheet1['Y4'].v.replace('Data da Consulta: ', ''))).toOFXDate();
  bankStatement.transactionId = aguid(bankStatement.dateTaken);
  bankStatement.accountId = aguid(workbook.Sheets.Sheet1['AB12'].v);

  while(workbook.Sheets.Sheet1['C' + index]){
  
    var transaction = {};
    transaction.date = date.addDays(workbook.Sheets.Sheet1['C' + index].v - 2);
    transaction.displayDate = transaction.date.toOFXDate();
    transaction.memo = workbook.Sheets.Sheet1['I' + index].v;
    transaction.amount = workbook.Sheets.Sheet1['Z' + index].v;
    transaction.balance = workbook.Sheets.Sheet1['AF' + index].v;
    transaction.id = aguid(transaction.date + transaction.memo + transaction.amount);
    transaction.type = transaction.amount >= 0 ? 'CREDIT' : 'DEBIT';
    transactions.push(transaction);
  
    index = index + 2;
  }

  var minimum = function(element, index, array){
    debugger;
    return array
      .map((arrayItem) => arrayItem >= element)
      .reduce((accumulator, value) => accumulator && value, true);
  }

  var maximum = function(element, index, array){
    return array
      .map((arrayItem) => arrayItem <= element)
      .reduce((accumulator, value) => accumulator && value, true);
  }

  bankStatement.transactions = transactions;
  lowestDate = transactions.map((transaction) => transaction.date).find(minimum);
  bankStatement.startDate = (new Date(lowestDate.getFullYear(), lowestDate.getMonth(), 1)).toOFXDate();
  highestDate = transactions.map((transaction) => transaction.date).find(maximum);
  bankStatement.endDate = (new Date(highestDate.getFullYear(), highestDate.getMonth() + 1, 0)).toOFXDate();
  bankStatement.balance = {};
  bankStatement.balance.amount = transactions[transactions.length - 1].balance;
  bankStatement.balance.dateAsOf = transactions[transactions.length - 1].date.toOFXDate();
  
  console.dir(transactions);
  
  fs.writeFile('./bank-statement.ofx', bankStatementTemplate(bankStatement));

});
