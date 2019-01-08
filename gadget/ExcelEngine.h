#ifndef EXCELENGINE_H
#define EXCELENGINE_H

#include <QObject>
#include <QAxObject>
#include <QList>
#include <QVariant>

#define TC_FREE(x)  {delete x; x=nullptr;}

class ExcelEngine : public QObject
{
    Q_OBJECT
public:
    explicit ExcelEngine(QObject *parent = nullptr);
    ExcelEngine(QString fileName);
    ~ExcelEngine();
    void Close();
    void ReadExeclData(QList<QList<QVariant> >& dataList);
    bool writeExcelData(QString savePath, QList<QList<QVariant> > &data);
    void convertToColName(int data, QString &res);
    QString to26AlphabetString(int data);

    bool writeToSheet(QList<QList<QVariant> > &data, QAxObject *worksheet);
    void castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res);
public slots:
    void open(QString fileName);
    void open();

signals:

private:
     QVariant ReadData();
signals:
     void openState(bool isSuccess);
     void readState(int status); //0 = 开始，1=结束
     void sigExcelData(QList<QList<QVariant> > dataList);

private:
    QString m_fileName = "";   //excel文件路径
    QString m_sheetName= "";
     QAxObject * m_excel;
     bool m_bIsOpen  = false;

     QAxObject*  m_books = nullptr;
     QAxObject*  m_book = nullptr;
     QAxObject*  m_sheets = nullptr;
     QAxObject*  m_sheet = nullptr;
};

#endif // EXCELENGINE_H
