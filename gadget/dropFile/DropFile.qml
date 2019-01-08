import QtQuick 2.9
import Qt.labs.platform 1.0
import QtQuick.Controls 2.2

Item {
    property alias text: statusText.text

    Button {
        anchors.left: parent.left
        anchors.leftMargin: 20
        anchors.verticalCenter: parent.verticalCenter

        //  width: parent.width/3
        //  height: parent.height
        // color:"red"
        Text {
            anchors.centerIn: parent
            text: "导出"
        }

        MouseArea {
            id: mouserII
            anchors.fill: parent
            hoverEnabled: true
            onClicked: {

                folderDialog.open()
            }
            onEntered: {
                // console.info("进来")
                mouserII.cursorShape = Qt.PointingHandCursor
            }
            onExited: {
                mouserII.cursorShape = Qt.ArrowCursor
            }
        }

        FolderDialog {
            id: folderDialog
            folder: StandardPaths.standardLocations(
                        StandardPaths.PicturesLocation)[0]
            onAccepted: {

                Excel.writeExcelData(folder)
            }
        }
    }
    Rectangle {
        id: fileRect
        width: parent.width / 3
        height: 100
        anchors.centerIn: parent
        color: "#DCDCDC"
        Text {
            id: text
            anchors.centerIn: parent
            text: qsTr("拖动或点击获取文件")
        }
        DropArea {
            anchors.fill: parent
            onDropped: {
                if (drop.hasUrls) {
                    for (var i = 0; i < drop.urls.length; i++) {
                        console.info('拖拽的文件url = ', drop.urls[i])
                        Excel.readExcel(drop.urls[i])
                    }
                }
            }
        }
        MouseArea {
            id: mouserArea
            anchors.fill: parent
            hoverEnabled: true
            onPressed: {
                //Excel.readExcel("C:/Users/fangd/Desktop/123.xls")
                fileDialog.open()
            }
            onEntered: {
                // console.info("进来")
                mouserArea.cursorShape = Qt.PointingHandCursor
            }
            onExited: {
                mouserArea.cursorShape = Qt.ArrowCursor
            }
        }
    }
    Item {
        anchors.left: fileRect.right
        width: parent.width / 3
        height: parent.height
        Text {
            id: statusText
            anchors.centerIn: parent
        }
    }

    FileDialog {
        id: fileDialog
        title: "选择一个Excel文件"
        folder: StandardPaths.writableLocation(StandardPaths.DocumentsLocation)
        nameFilters: ["Excel Files (*.xls  *.xlsx)", "*.*"]

        onAccepted: {
            console.log("You chose: " + fileDialog.file)
            Excel.readExcel(fileDialog.file)
        }
        onRejected: {
            console.log("Canceled")
        }
    }
}
