#ifndef EXCELJSONTABLE_H
#define EXCELJSONTABLE_H

// table : [ [][sheet0] [][sheet1] ... ]
#include "QXlsx/header/xlsxdocument.h"
#include <QObject>
#include <QJsonArray>
#include <QJsonObject>



class ExcelJsonTable : public QObject
{
    Q_OBJECT
public:
    explicit ExcelJsonTable(QJsonArray _sheetTitle, QJsonArray _table, QObject *parent = nullptr);
    void exportExcel(QString _outputPath);
    void writeCell(int sheetIndex, int row, int column, QJsonObject Obj);
    bool addSheet(QString sheetName);
    bool removeSheet(QString sheetName);
    QString getSheetName(int sheetIndex);
    int getStartRow(int sheetIndex);
    int getMaxColumnCount(int sheetIndex);
    int getSheetCount();

private:
    QJsonArray sheetTitle, tableArray;
    QString outputPath, sheetName;
    QXlsx::Document doc;

};

#endif // EXCELJSONTABLE_H
