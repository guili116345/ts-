#ifndef EXECLHELPER_H
#define EXECLHELPER_H

#include <QObject>
#include <QDebug>
#include <QFile>
#include <QDir>
#include <QAxObject>
#include "parsets.h"

class execlHelper : public QObject
{
    Q_OBJECT
public:
    explicit execlHelper(QObject *parent = nullptr);
    ~execlHelper();
    void newExecl(const QString &);
    void appendSheet(const QString &);
    void setCellValue(int row,int column,const QString &value); //按单元格写入
    void setValueByRow(const QString &,const QStringList &); //按行写入，未完成
    void setValueByCol(const QString &,const QStringList &); //按列写入
    void setValueByRange(QString start,QString end,QList<QList<QVariant>> cells);
    void saveExcel(const QString &fileName);

    /*
    *  brief  把列数转换为excel的字母列号
    *  param  data 大于0的数
    *  return 字母列号，如1->A 26->Z 27 AA
    */
    void convertToColName(int data, QString &res);
    /*
    *  brief 数字转换为26字母
    *  1->A 26->Z
    *  param data
    *  return
    */
    QString to26AlphabetString(int data);

    QAxObject * excel;
    QAxObject * workBooks;
    QAxObject * workBook;
    QAxObject * sheets;
    QAxObject * sheet;

signals:

public slots:

private:
    void list2ToQVariant(QList<QList<QVariant>>,QVariant &);
};

#endif // EXECLHELPER_H
