import QtQuick 2.9
import QtQuick.Window 2.2
import QtQuick.Controls 1.4

import "./dropFile/"

Window {
    id: window
    visible: true
    width: 640
    height: 480
    title: qsTr("Hello World")

    DropFile {
        id: dropFile
        width: window.width
        height: window.height / 3
    }
    ListModel {
        id: libraryModel
        //0
        ListElement {
            name: "总数量"
            PCT: "0"
        }
        //1
        ListElement {
            name: "净总数量"
            PCT: "0"
        }
        //2
        ListElement {
            name: "缺失数量"
            PCT: "0"
        }
        //3
        ListElement {
            name: "失效数量"
            PCT: "0"
        }
        //4
        ListElement {
            name: "none数量"
            PCT: "0"
        }
        //5
        ListElement {
            name: "正确数量"
            PCT: "0"
        }
        //6
        ListElement {
            name: "错误数量"
            PCT: "0"
        }
        //7
        ListElement {
            name: "正确率"
            PCT: "0%"
        }
        //8
        ListElement {
            name: "错误率"
            PCT: "0%"
        }
    }

    Rectangle {
        width: window.width
        height: window.height * 2 / 3
        anchors.top: dropFile.bottom
        border.color: "#00F5FF"
        TableView {
            id: tableview
            anchors.fill: parent
            TableViewColumn {
                horizontalAlignment: Text.AlignHCenter
                role: "name"
                title: "名称"
                width: tableview.width / 2
            }
            TableViewColumn {
                horizontalAlignment: Text.AlignHCenter
                role: "PCT"
                title: "个数"
                width: tableview.width / 2
            }

            headerDelegate: Rectangle {
                height: 50

                Text {
                    id: name
                    anchors.centerIn: parent
                    text: qsTr(styleData.value)
                }

                color: "#E0FFFF"
            }
            rowDelegate: Rectangle {
                height: 30
                color: styleData.selected ? "#7AC5CD" : (styleData.alternate ? "#DCDCDC" : "#FAEBD7")
            }
            itemDelegate: Rectangle {
                height: 30
                color: "transparent"
                //                border.color: "black"
                //                border.width: 1
                Text {
                    anchors.centerIn: parent

                    color: styleData.textColor

                    text: styleData.value
                }
            }

            model: libraryModel
        }
    }

    Connections {
        target: Excel
        onSummary: {

            console.info("获取到的ListFloatAAA", listFloat[1])
            listFloat[7] = listFloat[7] + "%"
            listFloat[8] = listFloat[8] + "%"
            console.info("获取到的ListFloatBBBB", listFloat)
            for (var i = 0; i < listFloat.length; i++)
                libraryModel.setProperty(i, "PCT", listFloat[i])
        }
        onUpdateTF: {
            libraryModel.setProperty(0, "PCT", count + "")
            libraryModel.setProperty(2, "PCT", inva + "")
            libraryModel.setProperty(3, "PCT", miss + "")
            libraryModel.setProperty(4, "PCT", none + "")
            libraryModel.setProperty(5, "PCT", t + "")
            libraryModel.setProperty(6, "PCT", f + "")
        }
        onStatusPro: {
            if (text == "1") {
                dropFile.text = ""
            } else {
                dropFile.text = dropFile.text + "\n" + text + "\n"
            }
        }
    }
}
