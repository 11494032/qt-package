#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QAxObject>
#include <QVariant>
#include <QDebug>
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}
void MainWindow::openExcel(QString fileName)
{
    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false);
    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open(const QString&)", fileName);

    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets

    int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目
    if (sheet_count > 0)
    {
        QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);
        QVariant var = readAll(work_sheet);
        castVariant2ListListVariant(var);
    }

    work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
    excel.dynamicCall("Quit(void)");  //退出
}

QVariant MainWindow::readAll(QAxObject *sheet)
{
    QVariant var;
    if (sheet != NULL && !sheet->isNull())
    {
        QAxObject *usedRange = sheet->querySubObject("UsedRange");
        if (NULL == usedRange || usedRange->isNull())
        {
            return var;
        }
        var = usedRange->dynamicCall("Value");
        delete usedRange;
    }
    return var;
}

void MainWindow::castVariant2ListListVariant(const QVariant &var)
{
    QVariantList varRows = var.toList();
    if (varRows.isEmpty())
    {
        return;
    }

    const int rowCount = varRows.size();
    QVariantList rowData;



    for (int i = 0; i < rowCount; ++i)
    {
        rowData = varRows[i].toList();

        if (i == 0)
        {
            QStringList headers;
            foreach (auto item, rowData)
            {
                QString value = item.toString();
                headers.append(value);
            }

            ui->tableWidget->setColumnCount(headers.size()); //设置列数
            ui->tableWidget->setHorizontalHeaderLabels(headers);
        }
        else
        {

            int row = ui->tableWidget->rowCount();
            ui->tableWidget->setRowCount(row + 1);
            for (int j = 0; j < rowData.size(); j++)
            {
                QString value = rowData[j].toString();
                QTableWidgetItem *item = new QTableWidgetItem(value);
                 ui->tableWidget->setItem(row, j, item);
            }
        }

    }

}

void MainWindow::on_pushButton_clicked()
{
   QString fileName = QFileDialog::getOpenFileName(this,tr("选择日志文件"),"",tr("(*.xlsx)")); //选择路径
   if (fileName.isEmpty())     //如果未选择文件便确认，即返回
       return;
   ui->lineEdit->setText(fileName);
   openExcel(fileName);

   //2018-02-11
   QString lastTime = ( ui->tableWidget->item(0,0)->text()).left(10);
   QString lastSeq = ui->tableWidget->item(0,1)->text();
   QString lastImei = ui->tableWidget->item(0,2)->text();
   qDebug()<<lastTime<<lastSeq<<lastImei;

   //时间相同，imei相同，序号一个个比较
   //循环读取数据
   for(int i=1;i<ui->tableWidget->rowCount();i++){
       // j<ui->tableWidget->columnCount()
       for(int j=0;j<3;j++){
           if(ui->tableWidget->item(i,j)!=NULL){               //一定要先判断非空，否则会报错
              // qDebug()<<ui->tableWidget->item(i,j)->text();
               if( j == 0 ){
                  if( ui->tableWidget->item(i,j)->text().startsWith(lastTime) ) { //同一天

                  }
                  else{
                    //写入登记表
                   QString lastTime = ( ui->tableWidget->item(0,0)->text()).left(10);
                  }
               }
           }
       }
   }

}
