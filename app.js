var express = require('express');
var multer = require('multer');
var morgan = require('morgan');
var upload = multer({dest:'./uploads/'});//where the file uploaded will be stored
var bodyParser = require('body-parser');
var Excel = require('exceljs'); //excelJS
var app = express();
app.use(morgan('dev'));
app.use(express.static(__dirname+'/client/app'));
app.use(bodyParser.json());
var _ = require('lodash');
var extractData = require('./test');
var workbook = new Excel.Workbook();

app.get('/',function(req, res, next){
	//res.status(200).end('All is fine');
})
app.post('/',upload.single('workbook'),function(req, res){
	res.status(200).end('Everything OK');
	//app.filename = req.file.filename;
	workbook.xlsx.readFile(__dirname+'/uploads/'+req.file.filename)
    .then(function() {
    	console.log(JSON.stringify(extractData(workbook)));
    });
});


var server = app.listen(8000, function(){
	console.log('app is running on: http://localhost:'+server.address().port);
});

