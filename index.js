const yahooFinance = require('yahoo-finance');
const XLSX = require('xlsx');
const Promise = require('bluebird');
const moment = require('moment');

let excelKey = {
	location: './../../../../GoogleDrive/stock.xlsx',
	tranSheet: 'transaction',
	checkSheet: 'dailyCheck'
};

let workbook = XLSX.readFile(excelKey.location);
let transactionData = XLSX.utils.sheet_to_json(workbook.Sheets[excelKey.tranSheet]);

let reduceTransaction = {};
let currentData = {
	totalInvestment: 0,
	potentialProfit: 0
}

transactionData.forEach(val=>{
	if(!reduceTransaction[val.Ticker]){
		reduceTransaction[val.Ticker] = {
			quantity: val.Quantity,
		}
		currentData.totalInvestment += val.Price * val.Quantity;
	} else {
		if(val.Action === 'Buy'){
			reduceTransaction[val.Ticker].quantity += val.Quantity;
			currentData.totalInvestment += val.Price * val.Quantity;
		} else if(val.Action === 'Sell'){
			reduceTransaction[val.Ticker].quantity -= val.Quantity;
			currentData.totalInvestment -= val.Price * val.Quantity;
		}
	}
});

let current = Promise.resolve()

for (let key in reduceTransaction) {
		current = current.then(()=>{
			console.log(key)
			return yahooFinance.quote({
				symbol: key,
				modules: [ 'price' ] // see the docs for the full list
			}, (err, quotes)=> {
				currentData.potentialProfit += reduceTransaction[key].quantity * quotes.price.regularMarketPrice;
				return
			});
		})
}

current.then(()=>{
	console.log(currentData)
	XLSX.utils.sheet_add_json(workbook.Sheets[excelKey.checkSheet]
	, [{
		"Date": moment().format('YYYYMMDD'),
		"TotalInvestment": currentData.totalInvestment,
		"PotentialProfit": currentData.potentialProfit
	}], {
		header: ['Date', 'TotalInvestment', "PotentialProfit"],
		skipHeader: true,
		origin: -1
	});
	XLSX.writeFile(workbook, excelKey.location);
});
