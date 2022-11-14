#ifndef PARSETS_H
#define PARSETS_H

#include <QXmlStreamReader>
#include <QFile>
#include <QMultiMap>
#include <QDebug>
#define O  qDebug() << "[" << __FILE__ << "]" <<  "[" << __FUNCTION__ << "]" <<  "[" << __LINE__ << "]: "

struct translate
{
    QString source;
    QString target;
    QString attribute;
};

class parseTS
{
public:
    parseTS();
    ~parseTS();

    bool readFile(const QString &fileName);
    int index;
    QMultiMap<int,QString> m_widgets;
    QMultiMap<int,QString> m_langes;
    QMultiMap<QString,translate> m_lan;


private:
    QXmlStreamReader reader;
    QString currentWidget;
    translate currentLange;
    void readTsElement();
    void readContextElement();
    void readNameElement();
    void readMessageElement();
    void readLocationElement();
    void readSourceElement();
    void readTranslationElement();

};

#endif // PARSETS_H
