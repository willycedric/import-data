var _ = require('lodash');
var extractDataFromFile = function(workbook){
			var worksheet = workbook.getWorksheet('Feuille 1');//get the sheet named 'Feuille 1'

			//Set the column contained in the file 
			worksheet.columns =[
				{header:'Id',key:'id',width:10},//A
				{header:'First Name',key:'firstName',width:32},//B
				{header:'Last Name',key:'lastName',width:32},//C
				{header:'Birth Date',key:'dob',width:43}//D
			];

			//TODO try to define those variables dynamically
			var firstNameCol = worksheet.getColumn('firstName');
			var lastNameCol=worksheet.getColumn('lastName');
			var dobCol = worksheet.getColumn('dob');
			var firstName=[];
			var lastName=[];
			var dob=[];
			var workbookObj =[];

			//get the data contained in each columns of the sheet
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
			//Discarding the column title (the first row of each column)
			firstName=_.without(firstName,firstNameCol.header);
			lastName =_.without(lastName, lastNameCol.header);
			dob =_.without(dob, dobCol.header);
			
			//building an object which represent the data contained in the workseet
			firstName.forEach(function(elt, index){
				workbookObj.push({firstName:elt,lastName:lastName[index], dob:dob[index]});
			});

			return workbookObj;
}


/*http://alexperry.io/node/2015/03/25/promises-in-node.html*/

module.exports = extractDataFromFile;