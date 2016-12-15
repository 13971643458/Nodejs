'use strict'
/*
 * 用法：
 * 1.配置环境：npm install node-xlsx
 * 2.命令行输入命令：node sweat.js 加班导出.xlsx
 * 3.运行成功后：当前目录会输出 加班统计.xlsx
 */

var xlsx 	= require("node-xlsx");
var fs 		= require('fs');
var inputFileName 	= process.argv[2];
var productName		='F项目'

class Sweat {

	constructor(inputFileName) {
		if (!inputFileName) {
			console.log('请输入xlsx文件.');
			process.exit(0);
		}
		this.inputFileName 		= inputFileName;
		this.outputFileName  	= '加班统计.xlsx'

		this.inputList			= [];
		this.outputList       	= [];
		this.outputList.push(['姓名','部门名称','打卡时间','进/出','备注','加班餐费','加班车费','市内因公外出']);
	}
	readXlsxFile(){
		var inputFileCache 		= xlsx.parse(this.inputFileName);
		this.inputList 			= inputFileCache[0].data;
	}
	formatXlsxFile(){

		let curDate,minDate,maxDate = null;

		for (let i = 1; i < this.inputList.length; i++) {

			let row 		= this.inputList[i];
			var date 		= new Date(1900, 0, row[3] - 1);
			let action 		= row[4];
			let isNeedInsterRow	= false;

			if (!curDate) {
				curDate = date;
				minDate = date;
				maxDate = date;
				console.log("拿当月第一天进minDate="+minDate);
			}
			if (this.isSameDay(curDate,date)) {
				if (action == '出') {
					maxDate = date;
				}
			}else{
				isNeedInsterRow = true;
			}

			if (isNeedInsterRow) {
				if (this.isSameDay(minDate,maxDate) && this.isAddedWork(maxDate)){
					this.outputList.push(this.crateInRow(row,minDate));
					this.outputList.push(this.crateOutRow(row,maxDate));
					console.log("拿当月出的数据maxDate="+maxDate);
				}
				curDate = date;
				minDate = date;
				// console.log("拿当月剩下来的进minDate="+minDate);
			}
		}

		let count = this.outputList.length;
		count = (count - 1)/2;
		console.log(count);
		let outRow 	= ['','','','','合计',''+25*count,'',''];
		this.outputList.push(outRow);
	}
	writeXlsxFile(){
		var xlsxObj		= [{name:'sheet1',data:this.outputList}];
		var file 		= xlsx.build(xlsxObj);

		fs.writeFileSync(this.outputFileName, file, 'binary');
	}


	/************ 辅助函数 ***********************/
	/*
	 *
	 */
	isSameDay(date1,date2){
		if (!date1 || !date2) {
			return true;
		}
		return (date1.getFullYear() == date2.getFullYear()) &&
			(date1.getMonth() == date2.getMonth()) &&
			(date1.getDate() == date2.getDate());

	}
	crateInRow(row,inTime){
		let inRow 	= ['姓名','部门','打卡时间','进',productName,'','',''];
		inRow[0] 	= row[1];
		inRow[1] 	= row[2];
		inRow[2] 	= inTime.toLocaleString();
		return inRow;
	}
	crateOutRow(row,outTime){
		let outRow 	= ['姓名','部门','打卡时间','出',productName,'25','',''];
		outRow[0] 	= row[1];
		outRow[1] 	= row[2];
		outRow[2] 	= outTime.toLocaleString();

		// 的士票
		if (outTime.getHours() >= 21) {
			outRow[6] = '的士票';
		}
		return outRow;
	}
	getDayCount(date){
		let day = String(date).split('.')[0]
		return day;
	}
	isAddedWork(date){
		return date.getHours() >= 20;
	}
}
var sweat = new Sweat(inputFileName);

sweat.readXlsxFile();
sweat.formatXlsxFile();
sweat.writeXlsxFile();
