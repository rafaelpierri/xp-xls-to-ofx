XLSX = require('xlsx');
aguid = require('aguid');
handlebars = require('handlebars');
fs = require('fs');

fs.readFile('./templates/bank-statement.hbs', 'utf8', function (err,data) {
  if (err) {
    return console.log(err);
  }

  var bankStatement = handlebars.compile(data);
  
  var workbook = XLSX.readFile('extrato.xls');
  
  Date.prototype.addDays=function(d){return new Date(this.valueOf()+864E5*d);};
  
  var date = new Date(Date.UTC(1900, 0, 1, 0, 0, 0, 0));
  
  var index = 25;
  var transactions = [];
  
  while(workbook.Sheets.Sheet1['C' + index]){
  
    var transaction = {};
    transaction.date = date.addDays(workbook.Sheets.Sheet1['C' + index].v - 2);
    transaction.memo = workbook.Sheets.Sheet1['I' + index].v;
    transaction.amount = workbook.Sheets.Sheet1['Z' + index].v;
    transaction.id = aguid(transaction.date + transaction.memo + transaction.amount);
    transaction.type = transaction.amount >= 0 ? 'CREDIT' : 'DEBIT';
    transactions.push(transaction);
  
    index = index + 2;
  }
  
  console.dir(transactions);
  
  fs.writeFile('./bank-statement.ofx', bankStatement({transactions: transactions}));

});
