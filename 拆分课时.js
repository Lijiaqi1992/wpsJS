/**
 * CommandButton1_Click Macro
 */
function 一键拆课时()
{
	for(var i = 1 ; i<=12; i++){
		Sheets.Item(i).Activate();
		work();
		格式调整2();
		边框();
		if(i == 2){
			Selection.Rows.Item(31).Delete();
			Selection.Rows.Item(32).Delete();
		}
		if(i==4 || i== 6 || i== 9 || i== 11){
			Selection.Rows.Item(32).Delete();
			Selection.Rows.Item(33).Delete();
		}
	}
}

/**
 * 格式调整2 Macro
 * 宏由 Administrator 录制，时间: 2023/12/21
 */
function 格式调整2()
{
	Columns.Item("A:B").Select();
	Selection.ColumnWidth = 7.370370;
	Columns.Item("C:K").Select();
	Selection.ColumnWidth = 21.814816;
	Columns.Item("L:M").Select();
	Selection.ColumnWidth = 3.851852;
	ActiveWindow.ScrollColumn = 10;
	Columns.Item("N:N").Select();
	Selection.ColumnWidth = 69.592590;
	ActiveWindow.ScrollColumn = 1;
	Columns.Item("C:K").Select();
	Selection.ColumnWidth = 19.592592;
	Rows.Item("2:18").Select();
	ActiveWindow.ScrollRow = 3;
	Rows.Item("2:19").Select();
	ActiveWindow.ScrollRow = 4;
	Rows.Item("2:20").Select();
	ActiveWindow.ScrollRow = 5;
	Rows.Item("2:21").Select();
	ActiveWindow.ScrollRow = 8;
	Rows.Item("2:27").Select();
	ActiveWindow.ScrollRow = 9;
	Rows.Item("2:28").Select();
	ActiveWindow.ScrollRow = 10;
	Rows.Item("2:29").Select();
	ActiveWindow.ScrollRow = 12;
	Rows.Item("2:30").Select();
	ActiveWindow.ScrollRow = 13;
	Rows.Item("2:31").Select();
	ActiveWindow.ScrollRow = 14;
	Rows.Item("2:32").Select();
	ActiveWindow.ScrollRow = 15;
	Rows.Item("2:33").Select();
	ActiveWindow.ScrollRow = 16;
	Rows.Item("2:37").Select();
	ActiveWindow.ScrollRow = 17;
	Rows.Item("2:33").Select();
	Selection.RowHeight = 70;

}
/**
 * 边框 Macro
 * 宏由 Administrator 录制，时间: 2023/12/21
 */
function 边框()
{
	ActiveWindow.ScrollRow = 29;
	Range("A1:N32").Select();
	(obj=>{
		obj.Weight = xlThin;
		obj.LineStyle = xlContinuous;
	})(Selection.Borders.Item(xlEdgeLeft));
	(obj=>{
		obj.Weight = xlThin;
		obj.LineStyle = xlContinuous;
	})(Selection.Borders.Item(xlEdgeTop));
	(obj=>{
		obj.Weight = xlThin;
		obj.LineStyle = xlContinuous;
	})(Selection.Borders.Item(xlEdgeBottom));
	(obj=>{
		obj.Weight = xlThin;
		obj.LineStyle = xlContinuous;
	})(Selection.Borders.Item(xlEdgeRight));
	(obj=>{
		obj.Weight = xlThin;
		obj.LineStyle = xlContinuous;
	})(Selection.Borders.Item(xlInsideVertical));
	(obj=>{
		obj.Weight = xlThin;
		obj.LineStyle = xlContinuous;
	})(Selection.Borders.Item(xlInsideHorizontal));
	Selection.Borders.Item(xlEdgeLeft).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlEdgeTop).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlEdgeBottom).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlEdgeRight).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlInsideVertical).ColorIndex = xlColorIndexAutomatic;
	Selection.Borders.Item(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic;

}




function work() {
	for (let k = 1; k <= 31; k++) {
		let Acol = Cells(k+1, 1);
		Acol.Value2=k;
	}
	for (var cell of Range("C2:K32")) {
		cell.Value2 = "NA";
		cell.Interior.Color = RGB(255, 255, 255);
		cell.Font.Size=10;
		cell.HorizontalAlignment = xlHAlignCenter;
	}
	let aList = [];
	let bList = [];
	let cList = [];
	let dList = [];
	let eList = [];
	let fList = [];
	let gList = [];
	let hList = [];
	let iList = [];
	let jList = [];
	let kList = [];
	let lList = [];
	let mList = [];
	let nList = [];
	let oList = [];
	let pList = [];
	for (var Ncol of Range("N2:N32")) {
		Ncol.Font.Size=10;
		let rowNum = Ncol.Row;

		let colNum = Ncol.Column;
		let Lcol = Cells(rowNum, 12);
		if (Ncol.Value2 != null) {

			let value = Ncol.Value2.split("\n");

			for (var i = 0; i < value.length; i++) {

				let resp = setValueAndColor(value[i], rowNum, (colNum + 1 + i));
				if(resp == "a"){
					aList.push("a");
				}else if(resp == "b"){
					bList.push("b");
				}else if(resp == "c"){
					cList.push("c");
				}else if(resp == "d"){
					dList.push("d");
				}else if(resp == "e"){
					eList.push("e");
				}else if(resp == "f"){
					fList.push("f");
				}else if(resp == "g"){
					gList.push("g");
				}else if(resp == "h"){
					hList.push("h");
				}else if(resp == "i"){
					iList.push("i");
				}else if(resp == "j"){
					jList.push("j");
				}else if(resp == "k"){
					kList.push("k");
				}else if(resp == "l"){
					lList.push("l");
				}else if(resp == "m"){
					mList.push("m");
				}else if(resp == "n"){
					nList.push("n");
				}else if(resp == "o"){
					oList.push("o");
				}else if(resp == "p"){
					pList.push("p");
				}
			}
			
			Lcol.Value2 = value.length
		}else{
			Lcol.Value2 = 0;
		}
	}
	Cells(36,3).Value2="1v1-白";
	Cells(36,3).Interior.Color = RGB(255, 255, 0);
	Cells(37,3).Value2=aList.length;
	Cells(36,4).Value2="1v1-黑";
	Cells(36,4).Interior.Color = RGB(146, 208, 80);
	Cells(37,4).Value2=cList.length;
	Cells(36,5).Value2="1v2-白";
	Cells(36,5).Interior.Color = RGB(255, 238, 173);
	Cells(37,5).Value2=bList.length;
	Cells(36,6).Value2="1v2-黑";
	Cells(36,6).Interior.Color = RGB(153, 221, 255);
	Cells(37,6).Value2=dList.length;
	
	Cells(38,3).Value2="1v1-白KEP/PET";
	Cells(38,3).Interior.Color = RGB(255, 156, 153);
	Cells(39,3).Value2=eList.length;
	Cells(38,4).Value2="1v1-黑KEP/PET";
	Cells(38,4).Interior.Color = RGB(158, 30, 26);
	Cells(38,4).Font.Color = RGB(255, 255, 255);
	Cells(39,4).Value2=gList.length;
	Cells(38,5).Value2="1v2-白KEP/PET";
	Cells(38,5).Interior.Color = RGB(255, 186, 132);
	Cells(39,5).Value2=fList.length;
	Cells(38,6).Value2="1v2-黑KEP/PET";
	Cells(38,6).Interior.Color = RGB(184, 96, 20);
	Cells(38,6).Font.Color = RGB(255, 255, 255);
	Cells(39,6).Value2=hList.length;
	
	Cells(40,3).Value2="Class-1v2";
	Cells(40,3).Interior.Color = RGB(154, 56, 215);
	Cells(40,3).Font.Color = RGB(255, 255, 255);
	Cells(40,3).Font.Bold = true;
	Cells(41,3).Value2=iList.length;
	Cells(40,4).Value2="Class-1v3";
	Cells(40,4).Interior.Color = RGB(245, 196, 0);
	Cells(41,4).Value2=jList.length;
	Cells(40,5).Value2="Class-1v4";
	Cells(40,5).Interior.Color = RGB(197, 202, 211);
	Cells(41,5).Value2=kList.length;
	Cells(40,6).Value2="Class-1v5";
	Cells(40,6).Interior.Color = RGB(255, 0, 0);
	Cells(40,6).Font.Color = RGB(255, 255, 255);
	Cells(40,6).Font.Bold = true;
	Cells(41,6).Value2=lList.length;
	
	Cells(42,3).Value2="2小时Class-1v2";
	Cells(42,3).Interior.Color = RGB(145, 156, 205);
	Cells(43,3).Value2=mList.length;
	Cells(42,4).Value2="2小时Class-1v3";
	Cells(42,4).Interior.Color = RGB(209, 131, 179)
	Cells(43,4).Value2=nList.length;
	Cells(42,5).Value2="2小时Class-1v4";
	Cells(42,5).Interior.Color = RGB(0, 255, 135);
	Cells(43,5).Value2=oList.length;
	Cells(42,6).Interior.Color = RGB(0,163,245);
	Cells(42,6).Font.Color = RGB(255, 0, 0);
	Cells(42,6).Value2="1小时1v1";
	Cells(42,6).Font.Bold = true;
	Cells(43,6).Value2=pList.length;
	
	Cells(42,7).Value2="当月总计";
	Cells(42,7).Font.Bold = true;
	Cells(43,7).Value2=aList.length+bList.length+cList.length+dList.length+eList.length+fList.length+gList.length+hList.length+iList.length+jList.length+kList.length
		+lList.length+mList.length+nList.length+oList.length+pList.length;
	
	Range("C36:G43").Select();
	Selection.HorizontalAlignment = xlHAlignCenter;
	Selection.Font.Size=10;
}

function setValueAndColor(value, rowNum, colNum) {

	var time = value.split("-")[0];
	
	if(time <= 0930){
		colNum = 3; //C列
	}else if(time > 0930 && time <=1100){
		colNum = 4;
	}else if(time > 1100 && time <=1230){
		colNum = 5;
	}else if(time > 1230 && time <=1400){
		colNum = 6; //F列
	}else if(time > 1400 && time <=1600){
		colNum = 7;
	}else if(time > 1600 && time <=1730){
		colNum = 8;
	}else if(time > 1730 && time <=1900){
		colNum = 9;
	}else if(time > 1900 && time <=1930){
		colNum = 10; //J列
	}else if(time > 1930 && time <=2100){
		colNum = 11; //K列
	}else {
		//>2100轮空，数据落在数据源（N列）后边
	}

	var wb = "白";

	if (time >= 1630) {

		wb = "黑";

	}

	var col = Cells(rowNum, colNum);
	col.Select();
	Selection.WrapText = true;
	col.Value2 = value;
	col.Font.Size=10;
	col.HorizontalAlignment = xlHAlignCenter;

	if ((value.match(/&/g) || []).length === 0) {
		if(value.indexOf("1小时") > -1){
			//1对1 1小时
			col.Interior.Color = RGB(0,163,245);
			col.Font.Color = RGB(255, 0, 0);
			col.Font.Bold = true;
			return "p";
		}
		
		if (value.indexOf("PET") > -1 || value.indexOf("KET") > -1) {

			//1对1 PET/KET
			wb == "白" ? col.Interior.Color = RGB(255, 156, 153) : petKetBlack1v1(col);
			return wb == "白"? "e":"g";

		} else {

			//1对1 普通
			wb == "白" ? col.Interior.Color = RGB(255, 255, 0) : col.Interior.Color = RGB(146, 208, 80);
			return wb == "白"?"a":"c";
		}

	} else if (value.indexOf("Class") < 0 && (value.match(/&/g) || []).length === 1) {

		if (value.indexOf("PET") > -1 || value.indexOf("KET") > -1) {

			//1对2 PET/KET
			wb == "白" ? col.Interior.Color = RGB(255, 186, 132) : petKetBlack1v2(col);
			return wb == "白"?"f":"h";
		} else {

			//1对2 普通
			wb == "白" ? col.Interior.Color = RGB(255, 238, 173) : col.Interior.Color = RGB(153, 221, 255);
			return wb == "白"?"b":"d";

		}

	} else if (value.indexOf("Class") > -1 
		&& value.indexOf("2小时") > -1 &&  (value.match(/&/g) || []).length === 3) {

		//班课2小时1v4
		col.Interior.Color = RGB(0, 255, 135)

		return "o";

	} else if (value.indexOf("Class") > -1 
		&& value.indexOf("2小时") > -1 &&  (value.match(/&/g) || []).length === 2) {

		//班课2小时1v3
		col.Interior.Color = RGB(209, 131, 179)

		return "n";

	} else if (value.indexOf("Class") > -1 
		&& value.indexOf("2小时") > -1 &&  (value.match(/&/g) || []).length === 1) {

		//班课2小时1v2
		col.Interior.Color = RGB(145, 156, 205)

		return "m";

	} else if (value.indexOf("Class") > -1 && (value.match(/&/g) || []).length === 1) {

		//班课1v2
		col.Interior.Color = RGB(154, 56, 215);
		
		col.Font.Color = RGB(255, 255, 255);
		col.Font.Bold = true;

		col.Value2 = value + "-【2人】";
		return "i";

	} else if (value.indexOf("Class") > -1 && (value.match(/&/g) || []).length === 2) {

		//班课1v3
		col.Interior.Color = RGB(245, 196, 0);

		col.Value2 = value + "-【3人】";

		col.Font.Color = RGB(255, 255, 255);
		col.Font.Bold = true;
		return "j";

	} else if (value.indexOf("Class") > -1 && (value.match(/&/g) || []).length === 3) {

		//班课1v4
		col.Interior.Color = RGB(197, 202, 211);

		col.Value2 = value + "-【4人】";
		return "k";

	} else if (value.indexOf("Class") > -1 && (value.match(/&/g) || []).length === 4) {

		//班课1v5
		col.Interior.Color = RGB(255, 0, 0);
		col.Font.Color = RGB(255, 255, 255);
		col.Font.Bold = true;
		col.Value2 = value + "-【5人】";
		return "l";

	}

}

//1对1 PET/KET 黑
function petKetBlack1v1(col) {

	col.Interior.Color = RGB(158, 30, 26);

	col.Font.Color = RGB(255, 255, 255);
	col.Font.Bold = true;

}

//1对2 PET/KET 黑
function petKetBlack1v2(col) {

	col.Interior.Color = RGB(184, 96, 20);

	col.Font.Color = RGB(255, 255, 255);
	
	col.Font.Bold = true;

}

