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
    explicit ExcelJsonTable(QJsonArray _titleArray, QJsonArray _tableArray, QString _outputPath, QList<int> _repeatedRows, double width = 1500, QObject *parent = nullptr);
    void exportExcel(bool _skipImages = true);
    void writeCell(int row, int column, QJsonObject Obj);
    void writeRow(QJsonArray Row);
    QString getSheetName(int row);
    int getStartRow(int sheetIndex);
    void updateColumnWidthMap(int sheetIndex);
    int getMaxColumnCount(int sheetIndex);
    int getSheetCount();

private:
    QJsonArray titleArray, tableArray;
    QString outputPath, sheetName;
    QList<int> repeatedRows;
    QMap<int, double> columnWidth;
    QXlsx::Document doc;
    int sheetIndex;
    double viewPortWidth;
    int currentRow, currentColumn;
    bool skipImages;

};

#endif // EXCELJSONTABLE_H
