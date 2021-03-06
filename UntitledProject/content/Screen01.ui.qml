/*
This is a UI file (.ui.qml) that is intended to be edited in Qt Design Studio only.
It is supposed to be strictly declarative and only uses a subset of QML. If you edit
this file manually, you might introduce QML code that is not supported by Qt Design Studio.
Check out https://doc.qt.io/qtcreator/creator-quick-ui-forms.html for details on .ui.qml files.
*/

import QtQuick
import QtQuick.Controls
import UntitledProject

Rectangle {
    width: 1280
    height: 800

    color: Constants.backgroundColor

    Button {
        id: btn_zamena_add
        x: 10
        y: 10
        width: 200
        height: 40
        text: qsTr("Создание замены")
        font.pointSize: 11
    }

    Button {
        id: btn_zamena_jurnal
        x: 220
        y: 10
        width: 200
        height: 40
        text: qsTr("Журнал замен")
        font.pointSize: 11
    }

    Button {
        id: btn_spisok_sotr
        x: 430
        y: 10
        width: 200
        height: 40
        text: qsTr("Список сотрудников")
        font.pointSize: 11
    }

    Button {
        id: btn_settings
        x: 640
        y: 10
        width: 200
        height: 40
        text: qsTr("Настройки")
        font.pointSize: 11
    }

    Frame {
        id: frame
        x: 10
        y: 60
        width: 1260
        height: 730
        layer.smooth: false
        rightPadding: 5
        bottomPadding: 5
        leftPadding: 5
        topPadding: 5
        wheelEnabled: false
        contentWidth: -1
    }
}

/*##^##
Designer {
    D{i:0;formeditorZoom:0.75}D{i:2}D{i:3}D{i:4}
}
##^##*/
