#include "exceljsontable.h"
#include <QJsonArray>
#include <QJsonObject>
#include <QDebug>

ExcelJsonTable::ExcelJsonTable(QJsonArray _titleArray, QJsonArray _tableArray, QString _outputPath, QList<int> _repeatedRows,double width, QObject *parent)
    : QObject{parent}
{
    tableArray = _tableArray;
    titleArray = _titleArray;

    outputPath = _outputPath;
    repeatedRows = _repeatedRows;
    viewPortWidth = width;
}

void ExcelJsonTable::exportExcel()
{
    QString sheetName = doc.currentSheet()->sheetName();
    QString newName = getSheetName(1);
    doc.renameSheet( sheetName, newName );
    sheetIndex = 0;
    int columnCount = getMaxColumnCount(sheetIndex);
    updateColumnWidthMap(sheetIndex);
    currentRow = 1;
    currentColumn = 1;

    QJsonArray  Row;
    QJsonObject obj;
    // write title
    for(int i=0; i < titleArray.size(); i++)
    {
        Row = titleArray[i].toArray();

    }

    //write table

    doc.saveAs(outputPath);
}

void ExcelJsonTable::writeCell(int sheetIndex, int row, int column, QJsonObject Obj)
{
    //doc.selectSheet(sheetIndex);
    //QString type = Obj["type"].toString();
    QVariant value = Obj.value("value").toVariant();
    QJsonObject style = Obj["style"].toObject();
    //name; width; height; color; background-color; font-family;  font-size; bold; align; border; row-span
    //QString name = style["name"].toString();
    QString color = style["color"].toString();
    QString backgroundColor = style["background-color"].toString();
    QString fontFamily = style["font-family"].toString();
    QString align =  style["align"].toString();
    double width = style["width"].toDouble();
    double height = style["height"].toDouble();
    int fontSize = style["font-size"].toInt();
    int border = style["border"].toInt();
    //int rowSpan = style["row-span"].toInt();
    bool bold = style["bold"].toBool();

    if(width == 0) width = 20;

    doc.setColumnWidth(column, width);
    doc.setRowHeight(row, height);
    QXlsx::Format format;
    QFont font(fontFamily);
    font.setBold(bold);
    font.setPointSize(fontSize);
    format.setFont(font);
    format.setFontColor(QColor(color));
    if(align.compare("center", Qt::CaseInsensitive) == 0)
            format.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    else if(align.compare("right", Qt::CaseInsensitive) == 0)
            format.setHorizontalAlignment(QXlsx::Format::AlignRight);
    else
            format.setHorizontalAlignment(QXlsx::Format::AlignLeft);
    format.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    format.setPatternBackgroundColor(QColor(backgroundColor));

    if(border > 2)
        format.setBorderStyle(QXlsx::Format::BorderThick);
    else if(border == 2)
            format.setBorderStyle(QXlsx::Format::BorderDouble);
    else if(border == 1)
        format.setBorderStyle(QXlsx::Format::BorderThin);
    else if(border == 0)
        format.setBorderStyle(QXlsx::Format::BorderNone);


    doc.write(row, column, value, format);
}

void ExcelJsonTable::writeCell(QJsonObject obj, QString cell)
{

}

void ExcelJsonTable::writeRow(QJsonArray Row, bool merge)
{
    if(merge)
    {

    }
    else
    {
        for(int i=0; i < Row.size(); i++)
        {

        }
    }
}

QString ExcelJsonTable::getSheetName(int row)
{
    if( tableArray[row].toArray().size() == 0 )
    {
        return QString("Sheet-").QString::number(row);
    }

    QJsonArray Row = tableArray[row].toArray();
    QJsonObject style = Row[0].toObject().value("style").toObject();
    return style["name"].toString();
}

int ExcelJsonTable::getStartRow(int sheetIndex)
{
    if(sheetIndex < 0) sheetIndex = 0;
    // find sheetIndex empty row
    if(sheetIndex == 0) return 0;

    int index = 0, temp = 0, row = 0;
    QJsonArray Row;
    for(int i=0; i < tableArray.count(); i++)
    {
        Row = tableArray[i].toArray();
        temp = Row.count();
        if(temp == 0) index++;
        if(index == sheetIndex)
        {
            row = i+1;
            break;
        }
    }
    return row;
}

void ExcelJsonTable::updateColumnWidthMap(int sheetIndex)
{
    int startRow = 0;
    if(sheetIndex != 0)
        startRow = getStartRow(sheetIndex);
    columnWidth.clear();
    double val;
    QJsonArray Row = tableArray[startRow].toArray();
    for(int i=0;i < Row.size(); i++)
    {
        val = Row[i].toObject().value("style").toObject().value("width").toDouble();
        columnWidth[i] = val;
    }
}

int ExcelJsonTable::getMaxColumnCount(int sheetIndex)
{
    int temp, count = 0;
    QJsonArray Row;
    int row = getStartRow(sheetIndex);
    for(int i=row; i < tableArray.count(); i++)
    {
        Row = tableArray[i].toArray();
        temp = Row.count();
        if(temp > count)
            count = temp;
        if(temp == 0) break;
    }

    return count;
}

int ExcelJsonTable::getSheetCount()
{
    // an empty row means new sheet
    int temp, count = 1;
    QJsonArray Row;
    for(int i=0; i < tableArray.count(); i++)
    {
        Row = tableArray[i].toArray();
        temp = Row.count();
        if(temp == 0)
            count++;
    }

    return count;
}

