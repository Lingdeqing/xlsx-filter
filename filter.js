const xlsx = require('xlsx')
const workbook = xlsx.readFile('1.xlsx')
const sheet = workbook.Sheets['Sheet2']
var rowCounter = 1
for(var r = 2; r <= 559; r++){
	var ok = false
	if(sheet['E'+r] && /南京/.test(sheet['E'+r].v)){
		ok = true
		rowCounter ++
	}

	for(var c = 65; c <= 79; c++){
		var _c = String.fromCharCode(c)
		if(ok && sheet[_c+r]){
			sheet[_c+rowCounter] = sheet[_c+r]
		}
		delete sheet[_c+r]
	}
}
sheet['!merges'] = []
xlsx.writeFile(workbook, '2.xlsx')