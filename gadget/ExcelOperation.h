#ifndef EXCELOPERATION_H
#define EXCELOPERATION_H

#include <QObject>
#include <QScopedPointer>
#include <QList>
#include <QVariant>
#include <QNetworkReply>
#include <QThread>
#include <QMap>
#include "ExcelBase.h"

class  ExcelEngine;

#define API  "http://47.105.196.123:10003/webQandA?"

class ExcelOperation : public QObject
{
    Q_OBJECT

public:
    QThread workerThread;
    explicit ExcelOperation(QObject *parent = nullptr);
    ExcelOperation(const QString & fileName);

    Q_INVOKABLE void readExcel(QString xlsFile);
    Q_INVOKABLE void networkRequest();


    void readExcelData(QString FileName);
    Q_INVOKABLE void writeExcelData(QString url);
private:
    void parseJson(QByteArray jsonData, QString question, int index);
signals:
    void summary(QList<QString> listFloat);
    void updateTF(int count, int t, int f, int inva, int miss, int none); //总数，正确，错误个数
    void statusPro(QString text);
    void startReadData(QList<QList<QVariant> >);

    void writeExcelData(QList<QList<QVariant> > cells);
public slots:
    void openstatusPro(bool status);
    void readStatus(int status);
    void setExcelData(QList<QList<QVariant> >);

    void stopExcelEngine();
private:
    QList< QList<QVariant> > m_datas;
    int m_count = 0; //总数量
    int m_Tcount = 0;   //成功个数
    int m_Fcount = 0;   //失败个数
    int m_invalidCount = 0; //无效的个数
    int m_missCount = 0; //缺失的个数
    int m_none = 0; //none的个数
    QNetworkReply * m_reply = nullptr;
    ExcelEngine * m_excelEngine = nullptr;
    QMap<int, QString> m_ForTMap;
    QMap<int , QString> m_failtMap;

     QScopedPointer<ExcelBase> m_xls;


};

#endif // EXCELOPERATION_H
