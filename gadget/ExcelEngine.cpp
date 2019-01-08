#include "ExcelEngine.h"
#include <QDebug>
#include <QTime>
#include <QVariantList>
#include "qt_windows.h"
#include "QMessageBox"

ExcelEngine::ExcelEngine(QObject *parent) : QObject(parent)
{
       HRESULT r = OleInitialize(0);
       if (r != S_OK && r != S_FALSE)
       {
           qDebug("Qt: Could not initialize OLE (error %x)", (unsigned int)r);
       }
}


ExcelEngine::ExcelEngine(QString fileName)
{
       HRESULT r = OleInitialize(nullptr);
       if (r != S_OK && r != S_FALSE)
       {
           qDebug("Qt: Could not initialize OLE (error %x)", (unsigned int)r);
       }

       m_fileName = fileName;
}

ExcelEngine::~ExcelEngine()
{
    if(m_excel)
    {
        if (!m_excel->isNull())
        {
            m_excel->dynamicCall("Quit()");
        }
    }
    TC_FREE(m_sheet);
    TC_FREE(m_sheets);
    TC_FREE(m_book);
    TC_FREE(m_books);
    TC_FREE(m_excel);

     OleUninitialize();
}

void ExcelEngine::open(QString fileName)
{
    if(m_bIsOpen)
    {
        Close();
    }

    m_excel = new QAxObject("Excel.Application", this);

   // m_excel->setProperty("Visible", false);
    if(m_excel)
    {
        m_bIsOpen = true;
        m_books = m_excel->querySubObject("WorkBooks");

        m_books->dynamicCall("Open (const QString&)", fileName);

        QVariant titleValue = m_excel->property("Caption");

        qDebug()<<QString("m_excel title: ")<<titleValue;
    }
    else
    {
        m_bIsOpen = false;
        Close();
    }

    emit openState(m_bIsOpen);


//        QAxObject * rows = usedRange->querySubObject("Rows");
//        QAxObject * columns = usedRange->querySubObject("Columns");
//        int rowStart = usedRange->property("Row").toInt();
//        int columnsStart = usedRange->property("Column").toInt();

//        int rowCount = rows->property("Count").toInt();
//        int columnCount = columns->property("Count").toInt();
//        qInfo()<<"获取到的rowCount"<<rowCount;
//        qInfo()<<"获取到的columnCount"<<columnCount;
//        time.start();
//        for(int i = rowStart ; i < rowCount; i++)
//        {
//            for(int j = columnsStart; j < columnCount; j++)
//            {
//                 QAxObject *cell = workItem->querySubObject("Cells(int, int)", i, j);
//                 QVariant  cellValue = cell->property("Value");
//                 QString message = QString("row-")+QString::number(i, 10)+QString("-column-")+QString::number(j, 10)+QString(":");
//                 qDebug()<<message<<cellValue;
//            }
//        }
//        qInfo()<<"读完的时间"<<time.elapsed()/1000.0<<"s";
}

void  ExcelEngine::open()
{
   open(m_fileName);
}

void ExcelEngine::Close()
{
    TC_FREE(m_sheet);
    TC_FREE(m_sheets);
    if (m_book != nullptr && ! m_book->isNull())
    {
        m_book->dynamicCall("Close(Boolean)", false);
    }
    TC_FREE(m_book);
    TC_FREE(m_books);
    if (m_excel != nullptr && !m_excel->isNull())
    {
        m_excel->dynamicCall("Quit()");
    }
    TC_FREE(m_excel);
    m_fileName  = "";
    m_sheetName = "";
     OleUninitialize();
}

QVariant  ExcelEngine::ReadData()
{
    QVariant cell;
    m_book = m_excel->querySubObject("ActiveWorkBook");

    m_sheets= m_book->querySubObject("Sheets");

    int sheetCount = m_sheets->property("Count").toInt();

    qDebug()<<QString("sheet count: ")<<sheetCount;

    for(int i = 1; i <= sheetCount;i++)
    {
        m_sheet = m_book->querySubObject("Sheets(int)",1);
        QString work_sheet_name  = m_sheet->property("Name").toString();
        QString message = QString("sheet ")+QString::number(1, 10)+QString(" name");

        qDebug()<<"message = "<<message<<"workSheetName"<<work_sheet_name;
    }

    if(sheetCount > 0 )
    {
        QTime time;

        m_sheet = m_book->querySubObject("Worksheets");
       // workSheet->setProperty("Activate");
        QAxObject *  workItem = m_sheet->querySubObject("Item(int)", 1);
        QAxObject * usedRange = workItem->querySubObject("UsedRange");
          time.start();

          cell = usedRange->dynamicCall("Value");

         qInfo()<<"读完的时间."<<time.elapsed()/1000.0<<"s";
         qInfo()<<"cell读取到的大小----"<<cell.toList().size();
    }

    return cell;
}

void ExcelEngine::ReadExeclData(QList<QList<QVariant> > & dataList)
{
    emit readState(0);
    QVariant excelVar = ReadData();
    QVariantList excelDataList = excelVar.toList();

    if(excelDataList.isEmpty())
    {
         return;
    }

    const int rowCount = excelDataList.size();
    QVariantList rowData;

    for(int i= 0; i < rowCount; i++)
    {
        rowData = excelDataList[i].toList();
        int EmptyCount = 0;
        for(auto data: rowData)
        {
                if(data.toString().isEmpty())
                {
                    EmptyCount++;
                }
        }
    //    qInfo()<<"EmptyCount"<<EmptyCount;
    //    qInfo()<<"rowData.length()"<<rowData.length();
        if(EmptyCount != rowData.length())
            dataList.append(rowData);
    }
    qInfo()<<"读取到的dataListd啊小"<<dataList.length();

    // Close();
//    emit sigExcelData(dataList);
//    emit readState(1);
}

bool  ExcelEngine::writeExcelData(QString savePath, QList<QList<QVariant> > &data)
{
    //QAxObject *work_books = m_excel->querySubObject("WorkBooks");

    //work_books->dynamicCall("Open(const QString&)", "D:\test.xls");

     //m_excel->setProperty("Caption", "Qt Excel");

    bool write_success = false;

     m_excel = new QAxObject("Excel.Application");
     //QAxObject excel();
     //excel.dynamicCall("SetVisible (bool Visible)", "false");
     //excel.setProperty("DisplayAlerts", false);

     m_books = m_excel->querySubObject("WorkBooks");
     if (!m_books)
     {
         QMessageBox::information(nullptr, "Export Error ", "No Microsoft Excel ", QMessageBox::Yes);
         return false;
     }
    m_books->dynamicCall("Add");

    m_book = m_excel->querySubObject("ActiveWorkBook");
    if (!m_book) return false;

    m_sheets = m_book->querySubObject("WorkSheets");
    if (!m_sheets) return false;
//    QAxObject * worksheet =  worksheets->querySubObject("Item(int)", 1);


//     QString sheetName = worksheet->property("Name").toString();

       //exportTofile
     //  int sheetsNum = worksheets->property("Count").toInt();

       m_sheet = m_sheets->querySubObject("Item(int)", 1);
       m_sheet->dynamicCall("Activate(void)");
       m_sheet->setProperty("Name", "QTExcelsheet");
//       write_success = writeToSheet(data, worksheet);


       int row = data.size();
       int col = data.at(0).size();
       QString rangStr;
       convertToColName(col, rangStr);
       rangStr += QString::number(row);
       rangStr = "A1:" + rangStr;
       qInfo()<<"rangStr = "<<rangStr;

       if (!m_sheet)
           return false;
       QAxObject* range = m_sheet->querySubObject("Range(const QString&)", rangStr);

       if (nullptr == range || range->isNull())
           return false;
       QVariant variant;

       QList< QList<QVariant> > datas;
       qInfo()<<"m_datasLength = "<<data.length();
       for(int i=0;i<data.length();++i)
       {
           QList<QVariant> rows;
           qInfo()<<"m_datas[i].length() = "<<i<<data[i].length();
           for(int j=0;j<data[i].length();++j)
           {

                   QVariant temp = data[i].at(j);
                   QString value = temp.toString().trimmed().simplified();
                   //QString value = "fafdfas 加价格";
                   qInfo()<<"value = "<<j<<value;
                   rows.append(value);


           }
           datas.append(rows);
       }

       castListListVariant2Variant(datas, variant);
       write_success = range->setProperty("Value", variant);

        delete range;

       //save
       savePath.replace('/', '\\');
       m_book->dynamicCall("SaveAs(const QString&)", savePath);
//       m_book->dynamicCall("Close(Boolean)", false);
//       m_excel->dynamicCall("Quit(void)");

       return write_success;
}

bool ExcelEngine::writeToSheet(QList<QList<QVariant>> &data, QAxObject *worksheet)
{
    int row = data.size();
    int col = data.at(0).size();
    QString rangStr;
    convertToColName(col, rangStr);
    rangStr += QString::number(row);
    rangStr = "A1:" + rangStr;
    qInfo()<<"rangStr = "<<rangStr;
    if (!worksheet)
        return false;
    QAxObject* range = worksheet->querySubObject("Range(QString)", rangStr);
    if (nullptr == range || range->isNull())
        return false;
    QVariant variant;
    castListListVariant2Variant(data, variant);
    range->setProperty("Value", variant);
  //   range->querySubObject("Value(QVariant)", variant);
   // range->dynamicCall("Value", var);
    //delete range;
    return true;
}

///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如1->A 26->Z 27 AA
///
void ExcelEngine::convertToColName(int data, QString &res)
{
    Q_ASSERT(data>0 && data<65535);
    int tempData = data / 26;
    if(tempData > 0)
    {
        int mode = data % 26;
        convertToColName(mode,res);
        convertToColName(tempData,res);
    }
    else
    {
        res=(to26AlphabetString(data)+res);
    }
}

///
/// \brief 数字转换为26字母
///
/// 1->A 26->Z
/// \param data
/// \return
///
QString ExcelEngine::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}

void ExcelEngine::castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res)
{
    QVariantList vars;
    const int rows = cells.size();
    for(int i=0;i<rows;++i)
    {
        vars.append(QVariant(cells[i]));
    }
    res = QVariant(vars);
}
