#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QApplication>
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow),
    saveExcelName(QApplication::applicationDirPath() + "result.xls")
{
    ui->setupUi(this);
    setWindowTitle(tr("XML Reader(cgtool_en.ts)"));
    parse.readFile(QApplication::applicationDirPath() + "/cgtool_en.ts");
}

MainWindow::~MainWindow()
{
    delete ui;
}

//void MainWindow::skipUnknownElement()
//{
//    reader.readNext();
//     while (!reader.atEnd())
//     {
//         if (reader.isEndElement())
//         {
//             reader.readNext();
//             break;
//         }

//         if (reader.isStartElement())
//         {
//             skipUnknownElement();
//         }
//         else
//         {
//             reader.readNext();
//         }
//     }

//}

void MainWindow::on_pushButton_clicked()
{
    m_execlHelper.newExecl(saveExcelName);
    int size = parse.m_langes.size();
    O << parse.m_langes.size();
    O << parse.m_widgets.size();
    O << parse.m_lan.size();
//    for(int i = 0 ;i < size;++i)
//    {
//        if(i%2 == 0)
//        {
//            m_execlHelper.setCellValue(i+1,1,parse.m_langes.value(i+1));
//        }
//        else
//        {
//            m_execlHelper.setCellValue(i,2,parse.m_langes.value(i+1));
//        }
//    }
    QStringList source ,target;
    for(int i = 0;i < parse.m_lan.size();i++)
    {
        source.append(parse.m_lan.values().at(i).source);
        target.append(parse.m_lan.values().at(i).target);
    }
    m_execlHelper.setValueByCol("A1",source);
    m_execlHelper.setValueByCol("B1",target);
//    QList<QList<QVariant>> tempText;
//    for(int i = 0;i < 3;i++)
//    {
//        QList<QVariant> data;
//        for(int j=0;j<2;j++)
//        {
//            data.append(QVariant(i*j));
//        }
//        tempText .append(data);
//    }
//    m_execlHelper.setValueByRange("A1","B3",tempText);
    m_execlHelper.saveExcel(saveExcelName);
    QMessageBox::information(this,QStringLiteral("提示"),QStringLiteral("数据已保存"));
}

void MainWindow::mysave()
{

    QString str = "记事本文件(*.text);;"
                  "c++(*.cpp);;"
                  "c source file(*.c);;"
                  "头文件(*.h);;"
                  "所有文件(*.*);;"
                  "java source file(*.java);;";
    QString defaultfile = "所有文件(*.*)";
    QString s =QFileDialog::getSaveFileName(this,"保存",QDir::currentPath(),str,&defaultfile);
    if(!s.isNull())
    {

//        QString text = this->text->toPlainText();
//        QFile f(s);
//        if(f.open(QIODevice::WriteOnly))
//        {
//           QByteArray arr;
//           arr.append(text);
//           f.write(arr);
//           f.close();
//        }
//        else
//        {
//            this->text->setText(f.errorString());}

//        this->setWindowTitle("mynote");
    }
}
