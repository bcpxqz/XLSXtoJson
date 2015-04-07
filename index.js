var XLSX = require('xlsx'),
	fs = require("fs"),
	path = require('path');

var en = {};
var zh = {};
var parentPath;



function readSheet(){
	var fileParh = 'json.xlsx';
	if(process.argv[2]){
		fileParh = process.argv[2];
	}
	if(!fs.existsSync(fileParh)){
		console.info('文件: '+fileParh+' 不存在');
	}
	parentPath = path.dirname(fileParh)+"/";
	var workbook = XLSX.readFile('json.xlsx'),
		sheetNames = workbook.SheetNames;
	for(var s=0; s< sheetNames.length; s++){
		var worksheet = workbook.Sheets[sheetNames[s]],
			region,
			mathes,
			startRow,
			endRow;
		if(worksheet){
			region = worksheet['!ref'];
			if(region){
				mathes = region.match(/([a-z]+)(\d+):([a-z]+)(\d+)/i);
				if(mathes){
					startRow = mathes[2]*1;
					endRow = mathes[4]*1;
					for(var r=startRow;r<=endRow;r++){
						if(worksheet['A'+r] && worksheet['B'+r] && worksheet['C'+r]){
							var key = worksheet['A'+r].v;
							en[key] = worksheet['B'+r].v.replace(/[\r\t\n]/g,"").trim();
							zh[key] = worksheet['C'+r].v.replace(/[\r\t\n]/g,"").trim();
						}
					}
				}
				
			}
		}
	}
}

function save(path){
	fs.writeFileSync(path+'en.json',JSON.stringify(en,null,4),{
		encoding:"utf-8"
	});
	fs.writeFileSync(path+'zh.json',JSON.stringify(zh,null,4),{
		encoding:"utf-8"
	});
}

function init(){
	readSheet();
	save(parentPath);
}
init();