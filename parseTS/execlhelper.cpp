#include "execlhelper.h"

execlHelper::execlHelper(QObject *parent) : QObject(parent)
{

}

execlHelper::~execlHelper()
{
    if(excel != Q_NULLPTR)
    {
        excel->dynamicCall("Quit()");
        delete excel;
        excel = Q_NULLPTR;
    }
    if(workBooks != Q_NULLPTR)
    {
        delete workBooks;
        workBooks = Q_NULLPTR;
    }
    if(workBook != Q_NULLPTR)
    {
        delete workBook;
        workBook = Q_NULLPTR;
    }
    if(sheets != Q_NULLPTR)
    {
        delete sheets;
        sheets = Q_NULLPTR;
    }
    if(sheet != Q_NULLPTR)
    {
        delete sheet;
        sheet = Q_NULLPTR;
    }
}

void execlHelper::newExecl(const QString &filename)
{
    QFile file(filename);
    excel = new QAxObject;
    excel->setControl("Excel.Application");
    excel->dynamicCall("SetVisible (bool)","false");//不显示窗体
    excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭时会出现类似“文件已修改，是否保存”的提示
    workBooks = excel->querySubObject("WorkBooks");//获取工作簿集合
    if(file.exists())
    {
        workBook = workBooks->querySubObject("open(const QString &)",filename);
    }
    else
    {
        workBooks->dynamicCall( "Add" );
        workBook = excel->querySubObject("ActiveWorkBook");
    }
    //默认有一个sheet
    sheets = workBook->querySubObject("Sheets");
    sheet = sheets->querySubObject("Item(int)",1);
}

void execlHelper::appendSheet(const QString & sheetName)
{
    int   cnt   =   1 ;
    QAxObject   *pLastSheet   =   sheets ->querySubObject( "Item(int)" ,   cnt);
    sheets ->querySubObject( "Add(QVariant)" ,   pLastSheet->asVariant());
    sheet   =   sheets ->querySubObject( "Item(int)" ,   cnt);
    pLastSheet->dynamicCall( "Move(QVariant)" ,   sheet ->asVariant());
    sheet ->setProperty( "Name", sheetName);
}

void execlHelper::setCellValue(int row, int column, const QString &value) // execl从1开始计数
{
    QAxObject *pRange = sheet->querySubObject("Cells(int,int)" ,row,column);
    //内容位置设置
    pRange->setProperty( "HorizontalAlignment" ,   -4131 );
    pRange->setProperty( "VerticalAlignment" ,   -4108 );
    pRange->dynamicCall( "SetValue",value);

    // pRange->setProperty("RowHeight", 50); //设置单元格行高
    //  pRange ->setProperty("ColumnWidth",   30);   //设置单元格列宽
    //  pRange ->setProperty("HorizontalAlignment",   -4108);   //左对齐（xlLeft）：-4131   居中（xlCenter）：-4108   右对齐（xlRight）：-4152
    //  pRange ->setProperty("VerticalAlignment",   -4108);   //上对齐（xlTop）-4160   居中（xlCenter）：-4108   下对齐（xlBottom）：-4107
    //  pRange ->setProperty("WrapText",   true);   //内容过多，自动换行
    //  pRange ->dynamicCall("ClearContents()");   //清空单元格内容
    //   QAxObject*   interior   =pRange ->querySubObject("Interior");
    //   interior->setProperty("Color",   QColor(0,   255,   0));   //设置单元格背景色（绿色）
    //   QAxObject*   border   =   cell->querySubObject("Borders");
    //   border->setProperty("Color",   QColor(0,   0,   255));   //设置单元格边框色（蓝色）
    //   QAxObject   *font   =   cell->querySubObject("Font");   //获取单元格字体 cell=worksheet->querySubObject("Cells(int,int)", 1, 3*i+2);
    //   font->setProperty("Name",   QStringLiteral("华文彩云"));   //设置单元格字体
    //   font->setProperty("Bold",   true);   //设置单元格字体加粗
    //   font->setProperty("Size",   20);   //设置单元格字体大小
    //   font->setProperty("Italic",   true);   //设置单元格字体斜体
    //   font->setProperty("Underline",   2);   //设置单元格下划线
    //   font->setProperty("Color",   QColor(255,   0,   0));   //设置单元格字体颜色（红色）
//    if (row   ==   1 ){
//    //加粗
//        QAxObject   *font   =   pRange->querySubObject( "Font" );   //获取单元格字体
//        font->setProperty( "Bold" , true );   //设置单元格字体加粗
    //    }
}

void execlHelper::setValueByRow(const QString & startPos,const QStringList &text) //未完成
{
    int size = text.size();
    if(size < 1)
    {
        O << "input text is empty!";
        return;
    }
    QString letter = startPos.left(1);
    int startNum = startPos.right(1).toInt();
    QString rangStr = startPos;
    rangStr += ":";
    rangStr += letter;
    rangStr += QString::number(startNum + size - 1);
    O <<rangStr;
    QAxObject *range = sheet->querySubObject("Range(const QString&)",rangStr);
    if(NULL == range || range->isNull())
    {
        O << "range is NULL!";
        return;
    }
    QList<QList<QVariant>> tempText;
    for(int i = 0;i < size;i++)
    {
        QList<QVariant> data;
        data.append(QVariant(text.at(i)));
        tempText .append(data);
    }
    QVariant vText;
    list2ToQVariant(tempText,vText);

    bool result = range->setProperty("Value2",vText);
    if(!result)
    {
        O << "range->setProperty is error!";
        return;
    }
}

void execlHelper::setValueByCol(const QString & startPos,const QStringList &text)
{
    int size = text.size();
    if(size < 1)
    {
        O << "input text is empty!";
        return;
    }
    if(startPos.size() != 2 && (startPos.at(0) < 'A' || startPos.at(0) > 'Z') && (startPos.at(1) < '1' || startPos.at(1) > '9'))
    {
        O << "input startPos is not right!";
        return;
    }
    QString letter = startPos.left(1);
    int startNum = startPos.right(1).toInt();
    QString rangStr = startPos;
    rangStr += ":";
    rangStr += letter;
    rangStr += QString::number(startNum + size - 1);
    O <<rangStr;
    QAxObject *range = sheet->querySubObject("Range(const QString&)",rangStr);
    range->setProperty("ColumnWidth",100);//设置单元格列宽
    range ->setProperty("WrapText",   true);   //内容过多，自动换行
    if(NULL == range || range->isNull())
    {
        O << "range is NULL!";
        return;
    }
    QList<QList<QVariant>> tempText;
    for(int i = 0;i < size;i++)
    {
        QList<QVariant> data;
        data.append(QVariant(text.at(i)));
        tempText .append(data);
    }
    QVariant vText;
    list2ToQVariant(tempText,vText);

    bool result = range->setProperty("Value2",vText);
    if(!result)
    {
        O << "range->setProperty is error!";
        return;
    }
}

void execlHelper::setValueByRange(QString start, QString end, QList<QList<QVariant>> cells)
{
    QString rangStr = start + ":" + end;
    O <<rangStr;
    QAxObject *range = sheet->querySubObject("Range(const QString&)",rangStr);
    if(NULL == range || range->isNull())
    {
        O << "range is NULL!";
        return;
    }
    QVariant vText;
    list2ToQVariant(cells,vText);
    bool result = range->setProperty("Value2",vText);
    if(!result)
    {
        O << "range->setProperty is error!";
        return;
    }
}

//保存Excel
void  execlHelper::saveExcel(const QString &fileName)
{
    workBook ->dynamicCall( "SaveAs(const QString &)" ,QDir::toNativeSeparators(fileName));
    workBook->dynamicCall("Close()");//关闭工作簿
    excel->dynamicCall("Quit()");//关闭excel
    delete excel;
    excel=Q_NULLPTR;
}

void execlHelper::list2ToQVariant(QList<QList<QVariant>> cells, QVariant &res)
{
    QVariantList vars;
    const int rows = cells.size();
    for(int i=0;i<rows;++i)
    {
        vars.append(QVariant(cells[i]));
    }
    res = QVariant(vars);
}

void execlHelper::convertToColName(int data, QString &res)
{
    Q_ASSERT(data>0 && data<65535);
    int tempData = data / 26;
    if(tempData > 0)
    {
        int mode = data % 26;
        convertToColName(mode,res);
        convertToColName(tempData,res);
    }
    else
    {
        res=(to26AlphabetString(data)+res);
    }
}

QString execlHelper::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}
