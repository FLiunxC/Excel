#include "NetworkRequest.h"
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonValue>
#include <QUrl>

NetworkRequest::NetworkRequest(QObject *parent) : QObject(parent)
{
    m_networkManager = new QNetworkAccessManager(this);
}

void NetworkRequest::addKeyValue(const QString &key, const QJsonValue &value)
{
    //构建json对象
    m_jsonObject.insert(key, value);
}

void NetworkRequest::addKeyValue(const QString &key, const QStringList &valuelist)
{
    //构建json数组对象
    QJsonArray jsonArray;

    for(QString value : valuelist)
    {
        jsonArray.append(value);
    }
    m_jsonObject.insert(key, jsonArray);
}

QNetworkReply *NetworkRequest::doPostRequest(const QString &targetUrl)
{
    QNetworkRequest request;
    request.setUrl(QUrl(targetUrl));

    return this->doPostRequest(request);
}

QNetworkReply* NetworkRequest::doPostRequest(const QNetworkRequest &request)
{
    QByteArray PostData;
    QNetworkRequest req = request;

#ifdef JSON
    req.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");
#else
    req.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
#endif

    // 构建 JSON 文档
    QJsonDocument document;

    document.setObject(m_jsonObject);

    PostData = document.toJson(QJsonDocument::Compact);

    QStringList keyList = m_jsonObject.keys();

    //将json清空
    foreach (QString it, keyList) {
        m_jsonObject.remove(it);
    }

    return m_networkManager->post(req,PostData);
}

QNetworkReply *NetworkRequest::doGetRequest(const QString &targetUrl)
{
    QNetworkRequest request;
    QString url = targetUrl;

    if(!url.endsWith("?"))
    {
        url = url +"?";
    }

    //get参数进行赋值
    QStringList keyList = m_jsonObject.keys();

    for(int i = 0; i < keyList.length(); i++)
    {
        QString key = keyList.at(i);
        QString argument = key +"="+ m_jsonObject.value(key).toString();
        if(i != keyList.length() - 1)
        {
            argument += "&";
        }

        url += argument;
    }

    //将json清空
    foreach (QString it, keyList) {
        m_jsonObject.remove(it);
    }
    qInfo()<<"请求到底 url = "<<url;
    request.setUrl(url);

    return this->doGetRequest(request);
}

QNetworkReply *NetworkRequest::doGetRequest(const QNetworkRequest &request)
{
    QNetworkRequest request2 = request;

    request2.setRawHeader("Content-Type","application/json");

    QNetworkReply *  reply = m_networkManager->get(request2);

    return reply;
}

QNetworkReply *NetworkRequest::uploadFile(const QString &Filepath, QString targetUrl)
{
    QString path = Filepath;
    QHttpMultiPart *multiPart = new QHttpMultiPart(QHttpMultiPart::FormDataType);
    QHttpPart imagePart;
    imagePart.setHeader(QNetworkRequest::ContentTypeHeader, QVariant("image/png"));
    imagePart.setHeader(QNetworkRequest::ContentDispositionHeader, QString("form-data; name=\"%1\"; filename=\"%2\"").arg("file").arg(path));
    QFile *file = new QFile(path);
    file->open(QIODevice::ReadOnly);
    imagePart.setBodyDevice(file);

    file->setParent(multiPart);
    multiPart->append(imagePart);

    QNetworkRequest request;
    request.setUrl(targetUrl);

    QNetworkReply *  reply = m_networkManager->post(request, multiPart);
    multiPart->setParent(reply);

    return reply;
}
