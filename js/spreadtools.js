function toolsClickEvents(spread) {
    $(".fa[class*='fa-']").click(function (e) {
        var operaType, className;
        className = e.target.className;
        //对齐，字体设置
        if (className.search('center|left|right|bold|italic') > -1) {
            operaType = 'cellStyle';
        } else if (className.search('table|columns') > -1) {
            operaType = 'cellOpera';
        }else if(className.search('deleterow|deletecolumn|addrow|addcolumn') > -1){
            operaType = 'rowAndcolumnOpera';
        }else if(className.search('eye')){
            operaType = 'preview';
        }
        eventExecute(spread, operaType, className);
    })
}

function eventExecute(spread, operaType, className) {
    var operaDetail;
    if (className.search('center') > -1) {
        operaDetail = 'center';
    } else if (className.search('left') > -1) {
        operaDetail = 'left';
    } else if (className.search('right') > -1) {
        operaDetail = 'right';
    } else if (className.search('table') > -1) {
        operaDetail = 'table';
    } else if (className.search('columns') > -1) {
        operaDetail = 'columns';
    }else if(className.search('bold') > -1){
        operaDetail = 'bold';
    }else if(className.search('italic') > -1){
        operaDetail = 'italic';
    }else if(className.search('deleterow') > -1){
        operaDetail = 'deleterow';
    }else if(className.search('deletecolumn') > -1){
        operaDetail = 'deletecolumn';
    }else if(className.search('addrow') > -1){
        operaDetail = 'addrow';
    }else if(className.search('addcolumn') > -1){
        operaDetail = 'addcolumn';
    }
    switch (operaType) {
        case 'cellStyle':
            cellStyleFormat(spread, operaDetail);
            break;
        case 'cellOpera':
            merOrsplitCell(spread, operaDetail);
            break;
        case 'rowAndcolumnOpera':
            rowAndcolumnOpera(spread,operaDetail);
            break;
        case 'preview':
            preview(spread);
            break;
    }
}

/**
 * 拆分或合并单元格
 */
function merOrsplitCell(spread, operaDetail) {
    var sheet = spread.getActiveSheet();
    var sel = sheet.getSelections();
    if (sel.length > 0) {
        sel = getActualCellRange(sel[sel.length - 1], sheet.getRowCount(), sheet.getColumnCount());
        switch (operaDetail) {
            case 'table':
                sheet.suspendPaint();
                for (var i = 0; i < sel.rowCount; i++) {
                    for (var j = 0; j < sel.colCount; j++) {
                        sheet.removeSpan(i + sel.row, j + sel.col);
                    }
                }
                sheet.resumePaint();
                break;
            case 'columns':
                sheet.addSpan(sel.row, sel.col, sel.rowCount, sel.colCount);
        }

    }
}

function getActualCellRange(cellRange, rowCount, columnCount) {
    if (cellRange.row == -1 && cellRange.col == -1) {
        return new spreadNS.Range(0, 0, rowCount, columnCount);
    }
    else if (cellRange.row == -1) {
        return new spreadNS.Range(0, cellRange.col, rowCount, cellRange.colCount);
    }
    else if (cellRange.col == -1) {
        return new spreadNS.Range(cellRange.row, 0, cellRange.rowCount, columnCount);
    }

    return cellRange;
}
/**
 * 鼠标悬浮选中效果
 */
function toolsActive() {
    $('.col-xs-1.menu i').bind('mouseover', function (e) {
        $(e.target).parent().addClass('active');
        // console.log(e);
    });
    $('.col-xs-1.menu').bind('mouseout', function (e) {
        $(e.target).parent().removeClass('active');
        //   console.log(e);
    });
    $('[data-toggle="tooltip"]').tooltip();
}
/**
 * 单元格样式设置
 * @param spread
 * @param cellStyle
 */
function cellStyleFormat(spread, cellStyle) {
    var sheet = spread.getActiveSheet();
    var sel = sheet.getSelections();
    if (sel.length > 0) {
        sel = getActualCellRange(sel[sel.length - 1], sheet.getRowCount(), sheet.getColumnCount());
        var col = sel.col;
        var row = sel.row;
        var colCount = sel.colCount;
        var rowCount = sel.rowCount;
        var cell;
        for (var i = col; i < col + colCount; i++) {
            for (var j = row; j < row + rowCount; j++) {
                cell = sheet.getCell(j, i);
                // cell.foreColor("red");
                // cell.backColor("yellow");
                switch (cellStyle) {
                    case 'center':
                        cell.vAlign(GC.Spread.Sheets.VerticalAlign.center);
                        cell.hAlign(GC.Spread.Sheets.HorizontalAlign.center);
                        break;
                    case 'left':
                        cell.vAlign(GC.Spread.Sheets.VerticalAlign.center);
                        cell.hAlign(GC.Spread.Sheets.HorizontalAlign.left);
                        break;
                    case 'right':
                        cell.vAlign(GC.Spread.Sheets.VerticalAlign.center);
                        cell.hAlign(GC.Spread.Sheets.HorizontalAlign.right);
                        break;
                    case 'bold':
                        var font = cell.font();
                        var cssStyle = sheet.getStyle(j,i)|| new GC.Spread.Sheets.Style();
                        if(font.search('bold') > -1){
                            font = font.replace('bold','');
                        }else{
                            font +=  ' bold';
                        }
                        cssStyle.font=font;
                        sheet.setStyle(j,i,cssStyle,GC.Spread.Sheets.SheetArea.viewport);
                        break;
                    case 'italic':
                        var font = cell.font();
                        var cssStyle = sheet.getStyle(j,i)|| new GC.Spread.Sheets.Style();
                        if(font.search('italic') > -1){
                            font = font.replace('italic','');
                        }else{
                            font =  'italic '+font;
                        }
                        cssStyle.font=font;
                        sheet.setStyle(j,i,cssStyle,GC.Spread.Sheets.SheetArea.viewport);
                        break;
                }
            }
        }
    }
}

/**
 * 行列操作
 * @param spread
 * @param operation
 */
function rowAndcolumnOpera(spread,operation) {
     var sheet = spread.getActiveSheet();
     var row = sheet.getActiveRowIndex();
     var column = sheet.getActiveColumnIndex();
    switch (operation){
        case 'deleterow':
            sheet.deleteRows(sheet.getActiveRowIndex(), sheet.getSelections()[0].rowCount);
            break;
        case 'deletecolumn':
            sheet.deleteColumns(sheet.getActiveColumnIndex(),sheet.getSelections()[0].colCount);
            break;
        case 'addrow':
            sheet.addRows(sheet.getActiveRowIndex(),1);
            break;
        case 'addcolumn':
            sheet.addColumns(sheet.getActiveColumnIndex(),1);
    }
}

/**
 * 获取数据输入区的边界
 * @param spread
 */
function getDataIndex(spread) {
    var sheet = spread.getActiveSheet();
    var rowCount = sheet.getRowCount();
    var columnCount = sheet.getColumnCount();
    var dataIndex = {};
    dataIndex.left = [0,0];
    dataIndex.right = [];
    //行边界
    for(var i = 0; i<rowCount; i++){
        for(var j = 0; j<columnCount;j++){
            if(sheet.getValue(i,j) != null){
                break;
            }
        }
        //循环完，没有break跳出，说明这一行全为空,
        if(j == columnCount){
            dataIndex.right[0] = i;
            break;
        }
    }
    //列边界
    for(var j = 0; j<columnCount; j++){
        for(var i = 0; i<rowCount;i++){
            if(sheet.getValue(i,j) != null){
                break;
            }
        }
        //循环完，没有break跳出，说明这一列全为空,
        if(i == rowCount){
            dataIndex.right[1] = j;
            break;
        }
    }
    return dataIndex;
}

/**
 * 预览
 * @param spread
 */
function preview(spread){
    var spJson = spread.toJSON();
    var dataIndex = getDataIndex(spread);
    //var spJson = JSON.stringify(spread.toJSON());
    $('#presp').css({'height':'400px'});
    var spread2 = new GC.Spread.Sheets.Workbook(document.getElementById('presp'), {sheetCount: 1});
    
    $( "#dialog-form" ).dialog({
        autoOpen: true,
        height: '600',
        width: '80%',
        modal: true,
        buttons: [{
            text: '提交',
            click: function() {
                var sp2Json = spread2.toJSON();
                var data = sp2Json.sheets.Sheet1.data;
                console.log(data);
                console.log(dataIndex);
                $( "#dialog-form" ).dialog("close");
                }
             },{
            text: '取消',
            click: function() {
               $( "#dialog-form").dialog("close");
            }
        }],
        open:function () {
            spread2.fromJSON(spJson);
            var sheet = spread2.getActiveSheet();
            sheet.setColumnCount(dataIndex.right[1],GC.Spread.Sheets.SheetArea.viewport);
            sheet.setRowCount(dataIndex.right[0]),GC.Spread.Sheets.SheetArea.viewport;
            sheet.options.isProtected = true;
            sheet.options.protectionOptions = {
                allowSelectLockedCells: true,
                allowSelectUnlockedCells: true,
                allowSort: false
            };
           // sheet.setText(5,5,"哈哈");
           // spread2.refresh();
          // spread2.print();
        },
        close: function() {
            spread2.destroy(); // form[ 0 ].reset();
           // allFields.removeClass( "ui-state-error" );
        }
    });
}

