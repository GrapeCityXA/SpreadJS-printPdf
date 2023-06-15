window.onload = function () {
    var spread = new spreadNS.Workbook(document.getElementById("ss"));
    document.getElementById('savePDF').onclick = function () {
        //下载pdf的方法
        spread.savePDF(
            function (blob) {
                //设置下载pdf的文件名
                saveAs(blob, 'download.pdf');
            },
            console.log,
            {
                title: 'Test Title',
                author: 'Test Author',
                subject: 'Test Subject',
                keywords: 'Test Keywords',
                creator: 'test Creator'
            });
    };
    var sheet = spread.getActiveSheet();
    sheet.suspendPaint();
    var style = new GC.Spread.Sheets.Style();
    //设置字体大小
    style.font = '15px Times';
    sheet.setDefaultStyle(style);
    //添加表格内容
    addTableContent(sheet);
    //添加饼图
    addPieContent(sheet);
    var printInfo = sheet.printInfo();
    //showBorder是否打印控件的外边框线
    printInfo.showBorder(true);
    //showGridLine是否打印网格线
    printInfo.showGridLine(true);
    //headerCenter是否打印表头中心
    printInfo.headerCenter("Family Business Costs");
    printInfo.headerLeft("&G");
    printInfo.footerCenter("&P/&N");

    // initSpread(spread);

    //打印的方法
    document.getElementById('btnPrint').onclick = function () {
        // used to adjust print range, should set with printInfo (refer custom print for detail)
        spread.sheets[0].setText(31, 8, " ");

        spread.print();
    };
    sheet.resumePaint();
};

var spreadNS = GC.Spread.Sheets;

//添加数据信息
function addTableContent (sheet) {

    sheet.addSpan(1, 0, 1, 7);
    //设置列高
    sheet.setRowHeight(1, 40);
    sheet.getCell(1, 0).value("Costs").font("28px Times").foreColor("#11114f").hAlign(spreadNS.HorizontalAlign.headerLeft).vAlign(spreadNS.VerticalAlign.center);
    //合并单元格
    sheet.addSpan(2, 0, 1, 7);
    sheet.setRowHeight(2, 30);
    //获取指定表单区域中的指定单元格
    sheet.getCell(2, 0).value("Family Business").font("18px Times").foreColor("#11114f").backColor("#f5f5f5").hAlign(spreadNS.HorizontalAlign.headerLeft).vAlign(spreadNS.VerticalAlign.center);
    sheet.setColumnWidth(0, 105);
    sheet.setRowHeight(3, 35);
    sheet.getCell(3, 0).value("Costs Elements").font("Bold 15px Times").foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.headerLeft).vAlign(spreadNS.VerticalAlign.center);
    sheet.setColumnWidth(1, 70);
    sheet.getCell(3, 1).value("2018").font("Bold 15px Times").foreColor("#171717").backColor("#dfe9fb").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.setColumnWidth(2, 70);
    sheet.getCell(3, 2).value("2019").font("Bold 15px Times").foreColor("#171717").backColor("#d1dffa").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.setColumnWidth(3, 70);
    sheet.getCell(3, 3).value("2020").font("Bold 15px Times").foreColor("#171717").backColor("#9bbaf3").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.setColumnWidth(4, 70);
    sheet.getCell(3, 4).value("2021").font("Bold 15px Times").foreColor("#ffffff").backColor("#5c7ee6").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.setColumnWidth(5, 70);
    sheet.getCell(3, 5).value("2022").font("Bold 15px Times").foreColor("#ffffff").backColor("#1346a4").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.setColumnWidth(6, 70);
    sheet.getCell(3, 6).value("Total").font("Bold 15px Times").foreColor("#ffffff").backColor("#102565").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    
    sheet.getCell(4, 0).value("Salaries").foreColor("#171717").backColor("#ffffff").vAlign(spreadNS.VerticalAlign.center);
    sheet.setValue(4, 1, 200);
    sheet.setValue(4, 2, 210);
    sheet.setValue(4, 3, 250);
    sheet.setValue(4, 4, 300);
    sheet.setValue(4, 5, 360);
    sheet.getCell(4, 1).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(4, 2).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(4, 3).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(4, 4).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(4, 5).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(4, 6).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    
    sheet.getCell(5, 0).value("Consulting").foreColor("#171717").backColor("#ffffff").vAlign(spreadNS.VerticalAlign.center);
    sheet.setValue(5, 1, 100);
    sheet.setValue(5, 2, 100);
    sheet.setValue(5, 3, 100);
    sheet.setValue(5, 4, 240);
    sheet.setValue(5, 5, 260);
    sheet.getCell(5, 1).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(5, 2).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(5, 3).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(5, 4).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(5, 5).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(5, 6).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
  
    sheet.getCell(6, 0).value("Hardware").foreColor("#171717").backColor("#ffffff").vAlign(spreadNS.VerticalAlign.center);
    sheet.setValue(6, 1, 2000);
    sheet.setValue(6, 2, 2500);
    sheet.setValue(6, 3, 0);
    sheet.setValue(6, 4, 3000);
    sheet.setValue(6, 5, 3200);
    sheet.getCell(6, 1).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(6, 2).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(6, 3).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(6, 4).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(6, 5).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(6, 6).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
   
    sheet.getCell(7, 0).value("Software").foreColor("#171717").backColor("#ffffff").vAlign(spreadNS.VerticalAlign.center);
    sheet.setValue(7, 1, 500);
    sheet.setValue(7, 2, 100);
    sheet.setValue(7, 3, 100);
    sheet.setValue(7, 4, 200);
    sheet.setValue(7, 5, 250);
    sheet.getCell(7, 1).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(7, 2).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(7, 3).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(7, 4).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(7, 5).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(7, 6).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    
    sheet.getCell(8, 0).value("Training").foreColor("#171717").backColor("#ffffff").vAlign(spreadNS.VerticalAlign.center);
    sheet.setValue(8, 1, 1100);
    sheet.setValue(8, 2, 1300);
    sheet.setValue(8, 3, 2500);
    sheet.setValue(8, 4, 1300);
    sheet.setValue(8, 5, 1250);
    sheet.getCell(8, 1).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(8, 2).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(8, 3).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(8, 4).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(8, 5).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(8, 6).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
   
    sheet.getCell(9, 0).value("Travel").foreColor("#171717").backColor("#ffffff").vAlign(spreadNS.VerticalAlign.center);
    sheet.setValue(9, 1, 2000);
    sheet.setValue(9, 2, 2000);
    sheet.setValue(9, 3, 1000);
    sheet.setValue(9, 4, 1500);
    sheet.setValue(9, 5, 0);
    sheet.getCell(9, 1).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(9, 2).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(9, 3).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(9, 4).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(9, 5).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(9, 6).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
   
    sheet.getCell(10, 0).value("Other").foreColor("#171717").backColor("#ffffff").vAlign(spreadNS.VerticalAlign.center);
    sheet.setValue(10, 1, 200);
    sheet.setValue(10, 2, 200);
    sheet.setValue(10, 3, 100);
    sheet.setValue(10, 4, 150);
    sheet.setValue(10, 5, 30);
    sheet.getCell(10, 1).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(10, 2).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(10, 3).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(10, 4).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(10, 5).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getCell(10, 6).foreColor("#171717").backColor("#ffffff").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
   
    sheet.setFormula(4, 6, "=SUM(B5:F5)");
    sheet.setFormula(5, 6, "=SUM(B6:F6)");
    sheet.setFormula(6, 6, "=SUM(B7:F7)");
    sheet.setFormula(7, 6, "=SUM(B8:F8)");
    sheet.setFormula(8, 6, "=SUM(B9:F9)");
    sheet.setFormula(9, 6, "=SUM(B10:F10)");
    sheet.setFormula(10, 6, "=SUM(B11:F11)");
    
    sheet.getRange(4,1,7,6).formatter("$0.0");
   
}

//添加饼状图的方法
function addPieContent(sheet) {
    //合并单元格
    sheet.addSpan(12, 0, 1, 4);
    //获取指定表单区域中的指定单元格
    sheet.getCell(12, 0).value("Total Costs").font("15px Times").foreColor("#11114f").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(13, 0, 9, 4);
    //在单元格中指定公式
    sheet.setFormula(13, 0, '=PIESPARKLINE(G5:G11,"#dfe9fb","#d1dffa","#9bbaf3","#5c7ee6","#1346a4","#102565", "#ededed")');
}

// function initSpread(spread) {
//     var sheet = spread.getActiveSheet();

//     sheet.suspendPaint();
//     sheet.options.allowCellOverflow = true;
//     sheet.name("Demo");

//     sheet.addSpan(1, 1, 1, 1);
//     sheet.setValue(1, 1, "Month");
//     sheet.setValue(2, 1, "January");
//     sheet.setValue(3, 1, "February");
//     sheet.setValue(4, 1, "March");
//     sheet.setValue(5, 1, "April");
//     sheet.setValue(6, 1, "May");
//     sheet.setValue(7, 1, "June");
//     sheet.setValue(8, 1, "July");
//     sheet.setValue(9, 1, "August");
//     sheet.setValue(10, 1, "September");
//     sheet.setValue(11, 1, "October");
//     sheet.setValue(12, 1, "November");
//     sheet.setValue(13, 1, "December");

//     sheet.addSpan(1, 2, 1, 1);
//     sheet.setValue(1, 2, "Quantity");
//     sheet.setValue(2, 2, 60);
//     sheet.setValue(3, 2, 100);
//     sheet.setValue(4, 2, 110);
//     sheet.setValue(5, 2, 90);
//     sheet.setValue(6, 2, 112);
//     sheet.setValue(7, 2, 137);
//     sheet.setValue(8, 2, 80);
//     sheet.setValue(9, 2, 88);
//     sheet.setValue(10, 2, 118);
//     sheet.setValue(11, 2, 122);
//     sheet.setValue(12, 2, 130);
//     sheet.setValue(13, 2, 135);

//     sheet.addSpan(1, 3, 1, 1);
//     sheet.setValue(1, 3, "Price");
//     sheet.setValue(2, 3, 30.0);
//     sheet.setValue(3, 3, 45.0);
//     sheet.setValue(4, 3, 10.0);
//     sheet.setValue(5, 3, 99.9);
//     sheet.setValue(6, 3, 89.0);
//     sheet.setValue(7, 3, 150.0);
//     sheet.setValue(8, 3, 35.0);
//     sheet.setValue(9, 3, 59.9);
//     sheet.setValue(10, 3, 47.9);
//     sheet.setValue(11, 3, 55.0);
//     sheet.setValue(12, 3, 32.0);
//     sheet.setValue(13, 3, 20.0);


//     sheet.addSpan(1, 4, 1, 1);
//     sheet.setValue(1, 4, "Total");
//     sheet.setFormula(2, 4, "=C3*D3");
//     sheet.setFormula(3, 4, "=C4*D4");
//     sheet.setFormula(4, 4, "=C5*D5");
//     sheet.setFormula(5, 4, "=C6*D6");
//     sheet.setFormula(6, 4, "=C7*D7");
//     sheet.setFormula(7, 4, "=C8*D8");
//     sheet.setFormula(8, 4, "=C9*D9");
//     sheet.setFormula(9, 4, "=C10*D10");
//     sheet.setFormula(10, 4, "=C11*D11");
//     sheet.setFormula(11, 4, "=C12*D12");
//     sheet.setFormula(12, 4, "=C13*D13");
//     sheet.setFormula(13, 4, "=C14*D14");

//     // set column width
//     sheet.setColumnWidth(0, 50);
//     sheet.setColumnWidth(1, 100);
//     sheet.setColumnWidth(2, 60);
//     sheet.setColumnWidth(3, 60);
//     sheet.setColumnWidth(4, 90);
//     sheet.setColumnWidth(5, 50);

//     //set alignment
//     sheet.getRange(1, 1, 1, 4).hAlign(GC.Spread.Sheets.HorizontalAlign.center)
//     //set formatter
//     sheet.getRange(2, 3, 12, 1).formatter("$ #,##0");
//     sheet.getRange(2, 4, 12, 1).formatter("$ #,##0");

//     //set color
//     sheet.getRange(1, 1, 1, 4).backColor("#69a569").foreColor("white");
//     //set lineBorder
//     sheet.getRange(2, 1, 12, 4).setBorder(new GC.Spread.Sheets.LineBorder("#A8A9AD", GC.Spread.Sheets.LineStyle.dotted), { all: true });

//     // addFigures(sheet);

//     sheet.resumePaint();
// }

// function addFigures(sheet) {

//     sheet.getRange(20, -1, 1, -1).visible(false);
//     sheet.addSpan(4, 6, 6, 3);
//     sheet.setFormula(4, 6, '=PIESPARKLINE(E3:E7, "#e6f0e6","#c9dec9","#a3c8a3","#73ab73","#d1ffd1")');
//     sheet.addSpan(10, 6, 1, 3);
//     sheet.getCell(10, 6).text("Figure 1").hAlign(1);
// }
   