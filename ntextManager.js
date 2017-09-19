var fs = require("fs");
var xl = require('excel4node');
var Converter = require("csvtojson").Converter;
var json2xls = require('json2xls');



var converter = new Converter({});
var wb = new xl.Workbook();
var clientName = process.argv[2];
var clientDir = "./" + clientName + "/";
var csvDir = clientDir+"csvFiles/";
var ntextDir = clientDir+"NTEXT/";
var fileArr = [];
var csvArr = [];
var fileTitle = 0;

fs.readdirSync(csvDir).forEach(file => {
	//console.log(file);
	fileArr.push(file);
});

//console.log(fileArr);

for(f = 0;f < fileArr.length; f++){
	var file = csvDir + fileArr[f];
	var csvEncoding = { encoding: 'utf16le' }; 
	var csvString = fs.readFileSync(file, csvEncoding).toString(); 
	
	csvArr.push(csvString);

	ConvertToJson(csvString).then(function(responseData) {
        ConvertToXls(responseData);
    }).catch(function(error) {
        console.log("ERROR: " + error)        
    });
	
}



function ConvertToJson(csvStr){
	//console.log("STEP: ConvertToJson");
	var testblah = csvStr.length;
	return new Promise(function(resolve,reject) {
		var jsonObj;
		var converter = new Converter({noheader:true});
		converter.fromString(csvStr, function(err,result){ 
	 		 // console.log(csvStr);
	 		 // console.log(result);
	 		 // console.log("result length: " + result.length);
	 		if(result.length > 0){
	 			jsonObj = result;	 			
	 			resolve(jsonObj);
	 		}else{
	 			console.log("ConvertToJson ERROR: " + err)
	 			reject(err);
	 		}	 		
	 	});
	})	
}

function ConvertToXls(jsonObj){
		//console.log("STEP: ConvertToXls");
		//console.log(jsonObj);
		var rows = jsonObj.length;		
		var columns = Object.keys(jsonObj[0]);

		//create new worksheet
		var ws = wb.addWorksheet(fileArr[fileTitle]);
		
		//rows
		for(r=0;r < rows;r++){
			//columns
			for(c=0;c < columns.length;c++){
				var thisCell = 	ws.cell(r+1,c+1);			
				var cellData = jsonObj[r][columns[c]];

				//.xlsx cells have a limit of 32,767 chars.  We need to create a .txt file for every NTEXT that exceeds this limit and add a reference to the cell.
				var primaryKey = jsonObj[r][columns[0]] + " - " + jsonObj[0][columns[c]];
				var dataLen = cellData.length
				if(dataLen){
					if(dataLen > 32767){
						console.log("DATA LIMIT EXCEEDED: " + dataLen);
						//fs.writeFileSync(ntextDir + fileArr[fileTitle],cellData, 'utf-8');
						
						fs.writeFileSync(CreateDir(fileArr[fileTitle],primaryKey),cellData, 'utf-8');
						thisCell.style({fill: {type: 'pattern', patternType: 'solid', fgColor: 'yellow'}});
						thisCell.string("Data limit exeeded for this cell.  See included file: " + ntextDir + fileArr[fileTitle] + "/" + primaryKey);
					}else{
						thisCell.string(cellData);
					}
				}			
			}
		}

	  	if(fileTitle == fileArr.length - 1){
	 		WriteFile();
	 	}else{
		 	fileTitle++;
		 }
	
}

function CreateDir(fileName,primKey){
	//console.log("KOMPLETE: " + ntextDir + fileName + "/" + primKey);

	//create directory for specific file
	if(!fs.existsSync(ntextDir + fileName)){
		fs.mkdirSync(ntextDir + fileName);
	}

	return ntextDir + fileName + "/" + primKey;	
}

function WriteFile(){
	console.log("STEP: WRITING DOCUMENT")
	wb.write(clientDir + clientName + '.xlsx');
}






 