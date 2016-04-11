//var Excel = require('exceljs');
//var data =require('./data');
var _ = require('lodash');


// construct a streaming XLSX workbook writer with styles and shared strings
/*var options = {
    filename: './streamed-workbook.xlsx',
    useStyles: true,
    useSharedStrings: true
};
var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
var worksheet = workbook.addWorksheet('Feuille 1','FF0000');

worksheet.columns =[
	{header:'Id',key:'id',width:10},//A
	{header:'First Name',key:'firstName',width:32},//B
	{header:'Last Name',key:'lastName',width:32},//C
	{header:'Birth Date',key:'dob',width:43}//D
];
data.family.forEach(function(elt){
	//console.log(elt);
	worksheet.addRow(elt);
});*/

var extractDataFromFile = function(workbook){
			var worksheet = workbook.addWorksheet('Feuille 1','FF0000');
			worksheet.columns =[
	{header:'Id',key:'id',width:10},//A
	{header:'First Name',key:'firstName',width:32},//B
	{header:'Last Name',key:'lastName',width:32},//C
	{header:'Birth Date',key:'dob',width:43}//D
];
			var firstNameCol = worksheet.getColumn('firstName');
			var lastNameCol=worksheet.getColumn('lastName');
			var dobCol = worksheet.getColumn('dob');
			var firstName=[];
			var lastName=[];
			var dob=[];
			var workbookObj =[];

			firstNameCol.eachCell(function(cell, rowNumber){
				if(rowNumber){
						firstName[rowNumber-1]=cell.value;
				}

			});
			lastNameCol.eachCell(function(cell,rowNumber){
				if(rowNumber){
							lastName[rowNumber-1]=cell.value;
				}
			});
			dobCol.eachCell(function(cell,rowNumber){
				if(rowNumber){
							dob[rowNumber-1]=cell.value;
				}
			});
			firstName=_.without(firstName,firstNameCol.header);
			lastName =_.without(lastName, lastNameCol.header);
			dob =_.without(dob, dobCol.header);
			firstName.forEach(function(elt, index){
				workbookObj.push({firstName:elt,lastName:lastName[index], dob:dob[index]});
			});

			console.log(JSON.stringify(workbookObj,2));

			//workbook.commit();
}


/*http://alexperry.io/node/2015/03/25/promises-in-node.html*/

module.exports = extractDataFromFile;