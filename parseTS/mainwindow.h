#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDebug>
#include <QFile>
#include <QMessageBox>
#include "parsets.h"
#include <QAxObject>
#include "execlhelper.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();


private slots:
    void on_pushButton_clicked();

private:

//    void skipUnknownElement();
    void mysave();
    parseTS parse;
    execlHelper m_execlHelper;
    QString saveExcelName;


private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
