#include <QGuiApplication>
#include <QQmlApplicationEngine>
#include "ExcelOperation.h"
#include <QQmlContext>


int main(int argc, char *argv[])
{
    QCoreApplication::setAttribute(Qt::AA_EnableHighDpiScaling);

    QGuiApplication app(argc, argv);

    QQmlApplicationEngine engine;
    ExcelOperation * excelOperation = new ExcelOperation();

    engine.rootContext()->setContextProperty("Excel", excelOperation);
    engine.load(QUrl(QStringLiteral("qrc:/main.qml")));
    if (engine.rootObjects().isEmpty())
        return -1;

    int ok =  app.exec();

    delete  excelOperation;

    return ok;
}
