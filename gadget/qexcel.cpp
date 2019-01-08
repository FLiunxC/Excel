#include <QAxObject>
#include <QAxWidget>
#include <QFile>
#include <QFileDialog>
#include <QDir>
#include <QStringList>
#include <QDebug>

#include "qexcel.h"

QExcel::QExcel(QString xlsFilePath, QObject *parent)
{
    excel = 0;
    workBooks = 0;
    workBook = 0;
    sheets = 0;
    sheet = 0;
//    CreateExcelFile(xlsFilePath);

    qDebug() << "123";
    excel = new QAxObject("Excel.Application", parent);
    qDebug() << "456";
    workBooks = excel->querySubObject("Workbooks");
    qDebug() << "789";
    QFile file(xlsFilePath);
    if (file.exists())
    {
        workBooks->dynamicCall("Open(const QString&)", xlsFilePath);
        workBook = excel->querySubObject("ActiveWorkBook");
        sheets = workBook->querySubObject("WorkSheets");
    }
}

QExcel::~QExcel()
{
    close();
}

void QExcel::close()
{
    excel->dynamicCall("Quit()");

    delete sheet;
    delete sheets;
    delete workBook;
    delete workBooks;
    delete excel;

    excel = 0;
    workBooks = 0;
    workBook = 0;
    sheets = 0;
    sheet = 0;
}

QAxObject *QExcel::getWorkBooks()
{
    return workBooks;
}

QAxObject *QExcel::getWorkBook()
{
    return workBook;
}

QAxObject *QExcel::getWorkSheets()
{
    return sheets;
}

QAxObject *QExcel::getWorkSheet()
{
    return sheet;
}

void QExcel::selectSheet(const QString& sheetName)
{
    sheet = sheets->querySubObject("Item(const QString&)", sheetName);
}

void QExcel::deleteSheet(const QString& sheetName)
{
    QAxObject * a = sheets->querySubObject("Item(const QString&)", sheetName);
    a->dynamicCall("delete");
}

void QExcel::deleteSheet(int sheetIndex)
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    a->dynamicCall("delete");
}

void QExcel::selectSheet(int sheetIndex)
{
    sheet = sheets->querySubObject("Item(int)", sheetIndex);
}

void QExcel::setCellString(int row, int column, const QString& value)
{
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);
    range->dynamicCall("SetValue(const QString&)", value);
}

void QExcel::setCellFontBold(int row, int column, bool isBold)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Bold", isBold);
}

void QExcel::setCellFontSize(int row, int column, int size)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Size", size);
}

void QExcel::mergeCells(const QString& cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

void QExcel::mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn)
{
    QString cell;
    cell.append(QChar(topLeftColumn - 1 + 'A'));
    cell.append(QString::number(topLeftRow));
    cell.append(":");
    cell.append(QChar(bottomRightColumn - 1 + 'A'));
    cell.append(QString::number(bottomRightRow));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

QVariant QExcel::getCellValue(int row, int column)
{
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);
    return range->property("Value");
}

void QExcel::save()
{
    workBook->dynamicCall("Save()");
    //workBook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators("F:/issachuang.xls"));
}

int QExcel::getSheetsCount()
{
    return sheets->property("Count").toInt();
}

QString QExcel::getSheetName()
{
    return sheet->property("Name").toString();
}

QString QExcel::getSheetName(int sheetIndex)
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    return a->property("Name").toString();
}

void QExcel::getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn)
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    *topLeftRow = usedRange->property("Row").toInt();
    *topLeftColumn = usedRange->property("Column").toInt();

    QAxObject *rows = usedRange->querySubObject("Rows");
    *bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;

    QAxObject *columns = usedRange->querySubObject("Columns");
    *bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
}

void QExcel::setColumnWidth(int column, int width)
{
    QString columnName;
    columnName.append(QChar(column - 1 + 'A'));
    columnName.append(":");
    columnName.append(QChar(column - 1 + 'A'));

    QAxObject * col = sheet->querySubObject("Columns(const QString&)", columnName);
    //col->setProperty("ColumnWidth", width);//这样设置没效果，必须要以下方式动态设置
    col->dynamicCall("SetColumnWidth(const int&)", width);
}

void QExcel::setCellTextCenter(int row, int column)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    //range->setProperty("HorizontalAlignment", -4108);//xlCenter//这样设置没效果，必须要以下方式动态设置
    range->dynamicCall("SetHorizontalAlignment(const int&)", -4108);
}

void QExcel::setCellTextWrap(int row, int column, bool isWrap)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("WrapText", isWrap);
}

void QExcel::setAutoFitRow(int row)
{
    QString rowsName;
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    QAxObject * rows = sheet->querySubObject("Rows(const QString &)", rowsName);
    rows->dynamicCall("AutoFit()");
}

void QExcel::insertSheet(QString sheetName)
{
    sheets->querySubObject("Add()");
    QAxObject * a = sheets->querySubObject("Item(int)", 1);
    a->setProperty("Name", sheetName);
}

void QExcel::mergeSerialSameCellsInAColumn(int column, int topRow)
{
    int a,b,c,rowsCount;
    getUsedRange(&a, &b, &rowsCount, &c);

    int aMergeStart = topRow, aMergeEnd = topRow + 1;

    QString value;
    while(aMergeEnd <= rowsCount)
    {
            value = getCellValue(aMergeStart, column).toString();
            while(value == getCellValue(aMergeEnd, column).toString())
            {
                    clearCell(aMergeEnd, column);
                    aMergeEnd++;
            }
            aMergeEnd--;
            mergeCells(aMergeStart, column, aMergeEnd, column);

            aMergeStart = aMergeEnd + 1;
            aMergeEnd = aMergeStart + 1;
    }
}

void QExcel::clearCell(int row, int column)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

void QExcel::clearCell(const QString& cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

int QExcel::getUsedRowsCount()
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    int topRow = usedRange->property("Row").toInt();
    QAxObject *rows = usedRange->querySubObject("Rows");
    int bottomRow = topRow + rows->property("Count").toInt() - 1;
    return bottomRow;
}

void QExcel::setCellString(const QString& cell, const QString& value)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("SetValue(const QString&)", value);
}

void QExcel::setCellFontSize(const QString &cell, int size)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Size", size);
}

void QExcel::setCellTextCenter(const QString &cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
//    range->setProperty("HorizontalAlignment", -4108);//xlCenter//这样设置没效果，必须要以下方式动态设置
    range->dynamicCall("SetHorizontalAlignment(const int&)", -4108);
}

void QExcel::setCellFontBold(const QString &cell, bool isBold)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
//    range->setProperty("Bold", isBold);//这样设置没效果，必须要以下方式动态设置
    range->dynamicCall("SetBold(bool)",isBold);
}

void QExcel::setCellTextWrap(const QString &cell, bool isWrap)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
//    range->setProperty("WrapText", isWrap);  //这样设置没效果，必须要以下方式动态设置
    range->dynamicCall("SetWrapText(bool)",isWrap);
}

void QExcel::setRowHeight(int row, int height)
{
    QString rowsName;
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    QAxObject * r = sheet->querySubObject("Rows(const QString &)", rowsName);
    //r->setProperty("RowHeight", height);
    r->dynamicCall("SetRowHeight(const int&)", height);
}

/***********************************************************
Function     :   CreateExcelFile
Description  :   新建一个Excel文件
Input        :   filePath 新建文件保存的路径包含文件名
Output       :   无
Return       :   无
Others       :
Author       :
History      :   1. 创建函数2014.12.31
***********************************************************/

//void QExcel::CreateExcelFile(QString filePath)
//{
//    //QString fileName = QFileDialog::getSaveFileName(NULL,"Save File",".","Excel File (*.xls)");
//    filePath.replace("/","\\");  //这一步很重要，c:/123.xls保存失败， 保存成功！
//    QFile file(filePath.append(".xls"));
//    if (!file.open(QIODevice::WriteOnly | QIODevice::Text))
//        return;
//    file.close();
//    if (filePath.isEmpty())
//    {
//        return;
//    }
//    QAxWidget _excel("Excel.Application");
//    _excel.setProperty("Visible",false);
//    QAxObject * _workbooks = _excel.querySubObject("WorkBooks");
//    _workbooks->dynamicCall("Add");   //添加一个新的工作薄
//    QAxObject * _workbook = _excel.querySubObject("ActiveWorkBook");
//     // _workbook->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",
//    //  QString("C:\\Users\\Administrator\\Desktop\\781.xls"),56,QString(""),QString(""),false,false);   //SaveAs有很多参数可选，具体请自己百度，这条语句是保存为OFFICE 2003格式
//    _workbook->dynamicCall("SaveAs (const QString&)", filePath);  //保存
//    _workbook->dynamicCall("Close (Boolean)", false);
//    _excel.dynamicCall("Quit (void)");
//}

