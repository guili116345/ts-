#include "parsets.h"
#include <QXmlStreamAttributes>

parseTS::parseTS():index(1)
{

}

parseTS::~parseTS()
{

}

bool parseTS::readFile(const QString &fileName)
{
    QFile file(fileName);
    if(!file.open(QFile::ReadOnly | QFile::Text))
    {
        O << "QFile::ReadOnly error!";
        return false;
    }
    O << "open file succcess";
    reader.setDevice(&file);
    while(!reader.atEnd())
    {
        if(reader.isStartElement())
        {
            if(reader.name() == "TS") //进入context元素
            {
                readTsElement();
            }
            else
            {
                O << reader.name();
                reader.raiseError("Not a valid context Element");
            }
        }
        else // '/n'
        {
            reader.readNext();
        }
    }

}

void parseTS::readTsElement()
{
    Q_ASSERT(reader.isStartElement() && reader.name() == "TS");
    reader.readNext();
    while(!reader.atEnd())
    {
        if(reader.isEndElement())
        {
            reader.readNext();
            break;
        }
        if(reader.isStartElement())
        {
            if(reader.name() == "context")
            {
                readContextElement();
            }
            else
            {
                O << reader.name();
                reader.raiseError("Not a valid context Element");
            }
        }
        else // '/n'
        {
            // O << reader.tokenString() << reader.text();
            reader.readNext();
        }
    }
}

void parseTS::readContextElement()
{
    Q_ASSERT(reader.isStartElement() && reader.name() == "context");
    reader.readNext();
    while(!reader.atEnd())
    {
        if(reader.isEndElement())
        {
            reader.readNext();
            break;
        }
        if(reader.isStartElement())
        {
            if(reader.name() == "name")
            {
                readNameElement();
            }
            else if(reader.name() == "message")
            {
                readMessageElement();
            }
            else
            {
                O << " it is not name or message ? ";
            }
        }
        else // '/n'
        {
            // O << reader.tokenString() << reader.text();
            reader.readNext();
        }
    }
}

void parseTS::readNameElement()
{
    Q_ASSERT(reader.isStartElement() && reader.name() == "name");
    QString name = reader.readElementText();
    if(reader.isEndElement())
    {
        qDebug() <<  QString::fromLocal8Bit("窗口: ") << name << endl;
        m_widgets.insertMulti(index,name);
        currentWidget = name;
        reader.readNext();
    }
    else
    {
        O << "name is not 1 line";
    }
}

void parseTS::readMessageElement()
{
    Q_ASSERT(reader.isStartElement() && reader.name() == "message");
    reader.readNext();
    while(!reader.atEnd())
    {
        if(reader.isEndElement())
        {
            reader.readNext();
            break;
        }
        if(reader.isStartElement())
        {
            if(reader.name() == "location")
            {
                readLocationElement();
            }
            else if(reader.name() == "source")
            {
                readSourceElement();
            }
            else if(reader.name() == "translation")
            {
                readTranslationElement();
            }
        }
        else
        {
            // O << reader.tokenString() << reader.text();
            reader.readNext();
        }
    }
}

void parseTS::readLocationElement()
{
    Q_ASSERT(reader.isStartElement() && reader.name() == "location");
    QXmlStreamAttributes attributes = reader.attributes();
    if(attributes.hasAttribute("filename"))
    {
        QString filename = attributes.value("filename").toString();
        if(attributes.hasAttribute("line"))
        {
            QString line = attributes.value("line").toString();
            qDebug() <<  QString::fromLocal8Bit("filename: ") << filename << "  " << QString::fromLocal8Bit("line: ") << line << endl;
        }
    }
    else
    {
        O << " it is not have attributes" << endl;
    }
    reader.readNext();
    if(reader.isEndElement())
    {

        reader.readNext();
    }

}

void parseTS::readSourceElement()
{
    Q_ASSERT(reader.isStartElement() && reader.name() == "source");
    QString source = reader.readElementText();
    qDebug() <<  QString::fromLocal8Bit("源文: ") << source << endl;
    m_langes.insertMulti(index++,source);
    currentLange.source = source;
    if (reader.isEndElement())
    {
        reader.readNext();
    }
}

void parseTS::readTranslationElement()
{
    Q_ASSERT(reader.isStartElement() && reader.name() == "translation");
    QXmlStreamAttributes attributes = reader.attributes();
    if(attributes.hasAttribute("type"))
    {
        QString type = attributes.value("type").toString();
        if(!type.compare("vanished")) //已屏蔽的翻译文本
        {
            currentLange.attribute = type;
        }
        else if(!type.compare("obsolete")) //过时的
        {
            currentLange.attribute = type;
        }
        else if(!type.compare("unfinished")) //还没有翻译的文本
        {
            currentLange.attribute = type;
        }
        else //未知的种类翻译文本
        {
             currentLange.attribute = "unknown";
        }
    }
    else //已翻译的文本
    {
        currentLange.attribute = "finished";
    }
    QString translation = reader.readElementText();
    qDebug() <<  QString::fromLocal8Bit("译文: ") << translation << endl;
    m_langes.insertMulti(index++,translation);
    currentLange.target = translation;
    m_lan.insertMulti(currentWidget,currentLange);
    if(reader.isEndElement())
    {
        reader.readNext();
    }

}
