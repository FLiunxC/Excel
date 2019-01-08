#include "ExcelOperation.h"
#include <QDebug>
#include <QElapsedTimer>
#include <QAxObject>
#include "ExcelEngine.h"
#include "NetworkRequest.h"
#include <QEventLoop>
#include <QJsonDocument>
#include <QJsonParseError>
#include <QJsonObject>
#include <QJsonValue>
#include <QMetaType>

ExcelOperation::ExcelOperation(QObject *parent) : QObject(parent)
{
    qRegisterMetaType<QList<QString> > ("QList<QString>");
    qRegisterMetaType<QList<QList<QVariant> > > ("QList<QList<QVariant> >");
}

void ExcelOperation::readExcel(QString xlsFile)
{

     readExcelData(xlsFile);
    // emit statusPro("1");
//    qDebug() << "获取到的xlsFile"<<xlsFile;
//    QString fileHead = "file:///";
//    emit statusPro("正在打开Excel...");
//    if(xlsFile.startsWith(fileHead))
//    {
//        xlsFile = xlsFile.remove(0, fileHead.length());
//        qDebug() << "解析后的 xlsFile"<<xlsFile;
//    }

//    if(m_excelEngine == nullptr)
//    {
//        m_excelEngine = new ExcelEngine();
//  //   m_excelEngine->moveToThread(&workerThread);
//    }


//    m_excelEngine->open(xlsFile);

//    m_excelEngine->ReadExeclData(m_datas);

//    delete m_excelEngine;
//    m_excelEngine = nullptr;


//    networkRequest();
//    connect(m_excelEngine, &ExcelEngine::sigExcelData, this, &ExcelOperation::setExcelData);
 //   connect(&workerThread, &QThread::finished, this, &ExcelOperation::stopExcelEngine);

//    connect(&workerThread, SIGNAL(started()), m_excelEngine, SLOT(open()));

//    connect(m_excelEngine, &ExcelEngine::readState, this, &ExcelOperation::readStatus);
//    connect(m_excelEngine, &ExcelEngine::openState, this, &ExcelOperation::openstatusPro);

//    connect(this, &ExcelOperation::startReadData, m_excelEngine, &ExcelEngine::ReadExeclData);
//    connect(this, &ExcelOperation::writeExcelData, m_excelEngine, &ExcelEngine::writeExcelData);

//    workerThread.start();
}

void ExcelOperation::networkRequest()
{
    m_Tcount = 0;
    m_Fcount = 0;
    m_count = 0;
    m_invalidCount = 0;
    m_missCount = 0;
    m_none = 0;
    NetworkRequest * network = new NetworkRequest(this);

    QTime time;
    time.start();
    m_count = m_datas.length();
    qInfo()<<"读取到的m_datas = "<<m_datas.length();
    for(int i = 1 ; i <m_count; i++)
    {
        qInfo()<<"length = "<<m_datas[i].length();
        if(m_datas[i].length() < 4)
        {
         //   statusPro(QString("Excel第%1行少于4列，跳过").arg(QString::number(i)));
            continue;
        }
        QString question =  m_datas[i].at(0).toString();;
        QString answer;
        QString type = m_datas[i].at(1).toString() ;
        if(type  == "COMPLETE_MATCH")
        {
             answer =  m_datas[i].at(2).toString();
        }
        else if(type  == "HIGHT_MATCH")
        {
            answer = m_datas[i].at(4).toString();
        }
//        else if(type  == "NONE_MATCH")
//        {
//            continue;
//        }

        question = "'"+question+"'";
        qInfo()<<"请求的问题："<<i<<question;

        network->addKeyValue("questions", question);
        m_reply = network->doGetRequest(API);

        QEventLoop eventloop;

        connect(m_reply, &QNetworkReply::finished, &eventloop, &QEventLoop::quit);

        eventloop.exec();
        QByteArray arrayJson =  m_reply->readAll();
        parseJson(arrayJson,  answer, i);
    }

    qInfo()<<"解析完毕时间:"<<time.elapsed()/1000.0;
    qInfo()<<"m_Tcount = "<<m_Tcount;
    qInfo()<<"m_Fcount = "<<m_Fcount;
    qInfo()<<"m_invalidCount = "<<m_invalidCount;
    qInfo()<<"m_missCount = "<<m_missCount;
    double Count = m_Tcount+m_Fcount;
    double T_T = double(m_Tcount / Count)*100;
    double f_f = 100 - T_T;

    QList<QString> strList{QString::number(m_datas.length()), QString::number(Count), QString::number(m_invalidCount), QString::number(m_missCount), QString::number(m_none), QString::number(m_Tcount),  QString::number(m_Fcount), QString::number(T_T,'g',4), QString::number(f_f, 'g',4)};

    qInfo()<<"floatList"<<strList;
    emit summary(strList);

    m_datas[0].append("是否正确");
    m_datas[0].append("返回正确答案");

     for(int i = 1; i < m_datas.length(); i++)
     {
         m_datas[i].append(m_ForTMap[i]);
         m_datas[i].append(m_failtMap[i]);
     }

//     qInfo()<<"请求完成"<<m_datas[0].length();
//     if(m_excelEngine == nullptr)
//     {
//         m_excelEngine = new ExcelEngine();
//   //   m_excelEngine->moveToThread(&workerThread);
//     }

//     m_excelEngine->writeExcelData("D:/Text.xls", m_datas);

//     delete m_excelEngine;
//     m_excelEngine = nullptr;

   //  writeExcelData();
    // emit writeExcelData(m_datas);

}

void ExcelOperation::readExcelData(QString FileName)
{
    if(m_xls.isNull())
        m_xls.reset(new ExcelBase);

    statusPro("准备打开Excel....");
    m_xls->open(FileName);
    statusPro("打开Excel完毕，准备读取...");

    m_xls->setCurrentSheet(1);


    m_xls->readAll(m_datas);
    statusPro("读取Excel完毕，等待分析...");

    networkRequest();
   statusPro("分析Excel完毕");
}

void ExcelOperation::writeExcelData(QString url)
{
    qInfo()<<"url = "<<url;
    QString xlsFile = url+"QtExcel.xls";
    QElapsedTimer timer;
    timer.start();
    if(m_xls.isNull())
        m_xls.reset(new ExcelBase);
    m_xls->create(xlsFile);
    QList< QList<QVariant> > datas;
    qInfo()<<"m_datasLength = "<<m_datas.length();
    for(int i=0;i<m_datas.length();++i)
    {
        QList<QVariant> rows;
        qInfo()<<"m_datas[i].length() = "<<i<<m_datas[i].length();
        for(int j=0;j<m_datas[i].length();++j)
        {
            if(j  != 3)
            {
                QVariant temp = m_datas[i].at(j);
                QString value = temp.toString().trimmed().simplified();
                //QString value = "fafdfas 加价格";
                qInfo()<<"value = "<<j<<value;
                rows.append(value);
            }

        }
        datas.append(rows);
    }
//    for(int i=0;i<1000;++i)
//    {
//        QList<QVariant> rows;
//        for(int j=0;j<100;++j)
//        {
//            QString a = "/;;）在吗\n";
//            rows.append(a);
//        }
//        datas.append(rows);
//    }
    qInfo()<<"datas = "<<datas.length();
    m_xls->setCurrentSheet(1);
    timer.restart();
    m_xls->writeCurrentSheet(datas);
    qDebug()<<"write cost:"<<timer.elapsed()<<"ms";timer.restart();
    m_xls->save();
}

void ExcelOperation::parseJson(QByteArray jsonData, QString questionAnswer, int index)
{
    qInfo()<<"请求获取结果"<<jsonData;
    QJsonParseError jsonError;

    QJsonDocument jsonDocument = QJsonDocument::fromJson(jsonData, &jsonError);

    if(!jsonDocument.isNull() && (jsonError.error == QJsonParseError::NoError))
    {
        if(jsonDocument.isObject())
        {
            QJsonObject jsonObject = jsonDocument.object();

            if(jsonObject.contains("reply"))
            {
                QJsonValue jsonValue = jsonObject.value("reply");
                if(jsonValue.isObject())
                {
                    jsonObject = jsonValue.toObject();
                    if(jsonObject.contains("answer"))
                    {
                        QJsonValue jsonValueAnswer = jsonObject.value("answer");
                        QString type = jsonObject.value("type").toString();
                        qInfo()<<"请求的Type: "<<type;
                        if(type == "T" || type == "F")
                        {
                            QString answer = jsonValueAnswer.toString();
                            qInfo()<<"请求的答案: "<<answer;
                            qInfo()<<"对应的答案: "<<questionAnswer;
                            m_failtMap[index] = answer;
                            if(questionAnswer.contains(answer))
                            {
                                m_Tcount++;
                                m_ForTMap[index] = "T";
                            }
                            else
                            {
                                m_Fcount++;
                                m_ForTMap[index] = "F";
                            }
                        }
                        else if(type != "none")
                        {
                            if(type == "无效")
                            {
                                m_invalidCount++;
                            }
                            else
                            {
                                m_missCount++;
                            }
                        }
                        else
                        {
                            m_none++;
                        }

                        emit updateTF(m_count, m_Tcount, m_Fcount, m_invalidCount,  m_missCount,m_none);
                    }
                }
            }
        }
    }
}

void ExcelOperation::openstatusPro(bool status)
{
    QString openStatus;
    if(status)
    {
        emit openStatus = "打开Excel成功";
        startReadData(m_datas);
    }
    else
    {
        openStatus = "打开Excel失败";
    }

    emit statusPro(openStatus);
}

void ExcelOperation::readStatus(int status)
{
    if(status == 0)
    {
        statusPro("准备读取Excel");
    }
    else
    {
        statusPro("读取Excel完毕");
        statusPro("解析数据中...");
      //  workerThread.quit();

        networkRequest();
        statusPro("解析数据完毕");
    }
}

void ExcelOperation::setExcelData(QList<QList<QVariant> > data)
{
    m_datas = data;

    // qInfo()<<"解析完毕---"<<m_datas;
}

void ExcelOperation::stopExcelEngine()
{
//    if(m_excelEngine != nullptr)
//    {
//        bool ok = workerThread.wait();
//        if(ok)
//        {
//            delete  m_excelEngine;
//            m_excelEngine = nullptr;
//        }
//    }
}
