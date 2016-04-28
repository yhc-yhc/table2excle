//写入excel操作
function writeToExl(excel){
	var excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:"+"excel"+"' xmlns='http://www.w3.org/TR/REC-html40'>";
    excelFile += "<head>";
    excelFile += "<!--[if gte mso 9]>";
    excelFile += "<xml>";
    excelFile += "<x:ExcelWorkbook>";
    excelFile += "<x:ExcelWorksheets>";
    excelFile += "<x:ExcelWorksheet>";
    excelFile += "<x:Name>";
    excelFile += "{worksheet}";
    excelFile += "</x:Name>";
    excelFile += "<x:WorksheetOptions>";
    excelFile += "<x:DisplayGridlines/>";
    excelFile += "</x:WorksheetOptions>";
    excelFile += "</x:ExcelWorksheet>";
    excelFile += "</x:ExcelWorksheets>";
    excelFile += "</x:ExcelWorkbook>";
    excelFile += "</xml>";
    excelFile += "<![endif]-->";
    excelFile += "</head>";
    excelFile += "<body>";
    excelFile += excel;
    excelFile += "</body>";
    excelFile += "</html>";
    var base64data = "base64," + $.base64({ data: excelFile, type: 0 });
    window.open('data:application/vnd.ms-'+"excel"+';filename=exportData.xls;' + base64data);
}
//转换元素的内容
function parseString(data){
    
     //content_data = data.html().trim();
     content_data = data.text().trim();
     
     //content_data = escape(content_data);
     
    return content_data;
 }
//根据结果集rs取到表头
function getTabH(rs){
	var theadHtml='<table>';
	theadHtml += '<thead>';
	if(rs.length==1){
		theadHtml += '<tr>';
		theadHtml += '<th colspan="6">'+'河南省高速公路联网公司机电系统维护'+'<i><u>'+rs[0].SYS_NAME+'</i></u>'+'设备日巡检表'+"</th>";
		theadHtml += '</tr></thead>';
		theadHtml += '<tbody>';
		theadHtml += '<tr>';
		theadHtml += '<td colspan="6"  align="left">'+'承包商：'+'<i><u>'+'江苏东南智能系统科技有限公司'+"</u></i></td>";
		theadHtml += '</tr>';
		theadHtml += '<tr>';
		theadHtml += '<td colspan="6"  align="left">'+'业主单位：'+'<i><u>'+'河南省高速公路联网监控收费通信服务有限公司'+"</u></i></td>";
		theadHtml += '</tr>';
		theadHtml += '<tr>';
		theadHtml += '<td colspan="4" align="left">'+'巡检日期:'+rs[0].CHK_TIME+"</td>";
		theadHtml += '<td colspan="2" align="left">'+'巡检人员:'+'<i><u>'+rs[0].USER_NAME+'</i></u>'+"</td>";
		theadHtml += '</tr></tbody></table>';
		return theadHtml;
	}else{
		return theadHtml;
	}
}
//根据结果集rs得到表格内容
function getTabB(rs){
	var tbodyHtml='<table border="1px">';
	tbodyHtml += '<thead><tr>';
	tbodyHtml += '<th colspan="5" bgcolor="#F1F1F1">'+'设备信息'+"</th>";
	tbodyHtml += '<th colspan="1" bgcolor="#F1F1F1">'+'设备状况'+"</th>";
	tbodyHtml += '</tr></thead><tbody>';
	$.each(rs[0].DVC_LIST,function(){
		tbodyHtml += '<tr>';
		tbodyHtml += '<td colspan="5" bgcolor="#F1F1F1">';
		tbodyHtml += '设备名称：'+this.TYPE_NAME+'&nbsp;||&nbsp;';
		tbodyHtml += '设备位置：'+this.PLACE_PATH+'&nbsp;||&nbsp;';
		tbodyHtml += '设备流水号：'+this.BQLS;
		tbodyHtml += "</td>";
		tbodyHtml += '<td bgcolor="#F1F1F1">';
		var content='';
		if(this.CHKITEMS.length>1){
			$.each(this.CHKITEMS,function() {
				content += this.CHKITEM_KEY;
				content += ':&nbsp;';
				content += this.CHKITEM_VALUE;
				content += '<br>';
			})
		}else{
			content=this.CHKITEMS[0].CHKITEM_VALUE;
		}
		tbodyHtml += content;
		tbodyHtml += "</td>";
		tbodyHtml += '</tr>';
	})
	tbodyHtml += '</tr></tbody></table>';
	return tbodyHtml;
}
//根据结果集rs取到表脚
function getTabF(rs){
	var tfootHtml='<table>';
	tfootHtml += '<tbody>';
	if(rs.length==1){
		tfootHtml += '<tr>';
		tfootHtml += '<td colspan="3"  align="left">'+'提交人:'+'<i><u>'+rs[0].USER_NAME+'</i></u>'+"</td>";
		tfootHtml += '<td colspan="3"  align="left">'+'审核人:'+'<i><u>'+rs[0].USER_NAME+'</i></u>'+"</td>";
		tfootHtml += '</tr></tbody></table>';
		return tfootHtml;
	}else{
		return tfootHtml;
	}
}
//根据结果集rs得到表格数据
function getDate(rs){
	var theadHtml = getTabH(rs);
	var tbodyHtml = getTabB(rs);
	var tfootHtml = getTabF(rs);
	var excel='';
	excel += theadHtml;
	excel += tbodyHtml;
	excel += tfootHtml;
	return excel;
	
}
//根据结果集rs导出excel
function toExcelBy(rs){
	var excel = getDate(rs);
	//alert(excel)
	writeToExl(excel);
}
