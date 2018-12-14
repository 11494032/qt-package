#include "mainwindow.h"
#include <QApplication>

#include "excelbase.h">

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();
    w.openExcel("f:py/imeis.xlsx");
    return a.exec();
}
