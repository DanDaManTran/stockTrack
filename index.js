const yahooFinance = require('yahoo-finance');
const XLSX = require('xlsx');
const Promise = require('bluebird');
const moment = require('moment');

let excelKey = {
	location: './stock.xlsx',
	tranSheet: 'transaction',
	checkSheet: 'dailyCheck'
};

let workbook = XLSX.readFile(excelKey.location);
let transactionData = XLSX.utils.sheet_to_json(workbook.Sheets[excelKey.tranSheet]);

let reduceTransaction = {};
let currentData = {
	totalInvestment: 0,
	potentialProfit: 0,
	StocktotalInvestment: 0,
	ETFtotalInvestment: 0,
	ETFBondtotalInvestment: 0,
	MFtotalInvestment: 0
}

transactionData.forEach(val=>{
	if(!reduceTransaction[val.Ticker]){
		reduceTransaction[val.Ticker] = {
			quantity: val.Quantity,
			type: val.Type
		}
		currentData[`${val.Type}totalInvestment`] += val.Price * val.Quantity;
		currentData[`${val.Type}potentialProfit`] = 0;
		currentData.totalInvestment += val.Price * val.Quantity;
	} else {
		if(val.Action === 'Buy'){
			reduceTransaction[val.Ticker].quantity += val.Quantity;
			currentData[`${val.Type}totalInvestment`] += val.Price * val.Quantity;
			currentData.totalInvestment += val.Price * val.Quantity;

		} else if(val.Action === 'Sell'){
			reduceTransaction[val.Ticker].quantity -= val.Quantity;
			currentData[`${val.Type}totalInvestment`] -= val.Price * val.Quantity;
			currentData.totalInvestment -= val.Price * val.Quantity;
		}
	}
});

let current = Promise.resolve()

for (let key in reduceTransaction) {
		current = current.then(()=>{
			return yahooFinance.quote({
				symbol: key,
				modules: [ 'price' ] // see the docs for the full list
			}, (err, quotes)=> {

				currentData[`${reduceTransaction[key].type}potentialProfit`] += reduceTransaction[key].quantity * quotes.price.regularMarketPrice;
				currentData.potentialProfit += reduceTransaction[key].quantity * quotes.price.regularMarketPrice;
				return
			});
		})
}

current.then(()=>{
	XLSX.utils.sheet_add_json(workbook.Sheets[excelKey.checkSheet]
	, [{
		"Date": moment().format('YYYYMMDD'),
		"TotalInvestment": currentData.totalInvestment.toFixed(2),
		"PotentialProfit": currentData.potentialProfit.toFixed(2),
		"StocktotalInvestment": currentData.StocktotalInvestment.toFixed(2),
		"StockpotentialProfit": currentData.StockpotentialProfit.toFixed(2),
		"ETFtotalInvestment": currentData.ETFtotalInvestment.toFixed(2),
		"ETFpotentialProfit": currentData.ETFpotentialProfit.toFixed(2),
		"ETFBondtotalInvestment": currentData.ETFBondtotalInvestment.toFixed(2),
		"ETFBondpotentialProfit": currentData.ETFBondpotentialProfit.toFixed(2),
		"MFtotalInvestment": currentData.MFtotalInvestment.toFixed(2),
		"MFpotentialProfit": currentData.MFpotentialProfit.toFixed(2)
	}], {
		header: ['Date', 'TotalInvestment', "PotentialProfit", "StocktotalInvestment", "StockpotentialProfit", "ETFtotalInvestment", "ETFpotentialProfit", "ETFBondtotalInvestment", "ETFBondpotentialProfit", "MFtotalInvestment", "MFpotentialProfit"],
		skipHeader: true,
		origin: -1
	});
	
	XLSX.writeFile(workbook, excelKey.location);
});
