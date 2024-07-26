#ifndef EXCELJSONTABLE_H
#define EXCELJSONTABLE_H

#include <QObject>
#include <QJsonArray>
#include <QJsonObject>



class ExcelJsonTable : public QObject
{
    Q_OBJECT
public:
    explicit ExcelJsonTable(QObject *parent = nullptr);

private:
    QJsonArray header, table;

};

#endif // EXCELJSONTABLE_H
