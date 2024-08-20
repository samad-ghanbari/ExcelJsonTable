#include "exceljsontable.h"
#include <QJsonArray>
#include <QJsonObject>
#include <QDebug>
#include <QtMath>

ExcelJsonTable::ExcelJsonTable(QJsonArray _titleArray, QJsonArray _tableArray, QString _outputPath, QList<int> _repeatedRows,double width, QObject *parent)
    : QObject{parent}
{
    tableArray = _tableArray;
    titleArray = _titleArray;

    outputPath = _outputPath;
    repeatedRows = _repeatedRows;
    viewPortWidth = width;
}

void ExcelJsonTable::exportExcel(bool _skipImages)
{
    QString sheetName = doc.currentSheet()->sheetName();
    QString newName = getSheetName(1);
    doc.renameSheet( sheetName, newName );
    skipImages = _skipImages;
    sheetIndex = 0;
    doc.selectSheet(sheetIndex);
    updateColumnWidthMap(sheetIndex);
    currentRow = 1;
    currentColumn = 1;

    QJsonArray  Row;
    QJsonObject obj;
    // write title
    for(int i=0; i < titleArray.size(); i++)
    {
        Row = titleArray[i].toArray();
        writeRow(Row);
    }

    int columnCount = getMaxColumnCount(sheetIndex);
    QXlsx::CellRange range(currentRow, 1, currentRow, columnCount);
    QXlsx::Format format;
    format.setPatternBackgroundColor(QColor(QString("#55A")));
    doc.setRowHeight(currentRow,10);
    doc.mergeCells(range, format);
    currentRow++;

    //write table
    for(int i=0; i < tableArray.size(); i++)
    {
        Row = tableArray[i].toArray();
        writeRow(Row);
    }


    doc.saveAs(outputPath);
}

void ExcelJsonTable::writeCell(int row, int column, QJsonObject Obj)
{
    // cell index starts from 1
    // sheet index starts from 0

    //doc.selectSheet(sheetIndex);
    QString type = Obj["type"].toString();
    if(type.compare("img", Qt::CaseInsensitive) == 0)
        if(skipImages) return;

    QVariant value = Obj.value("value").toVariant();
    QJsonObject style = Obj["style"].toObject();
    //name; width; height; color; background-color; font-family;  font-size; bold; align; border; row-span
    //QString name = style["name"].toString();
    QString color = style["color"].toString();
    QString backgroundColor = style["background-color"].toString();
    QString fontFamily = style["font-family"].toString();
    QString align =  style["align"].toString();
    // 1cm = 37.80pt
    // 1pt = 72/96 px = 0.75 px
    double width = style["width"].toDouble() /37.80;
    double height = style["height"].toDouble();
    int fontSize = style["font-size"].toInt();
    int border = 1;//style["border"].toInt();
    //int rowSpan = style["row-span"].toInt();
    bool bold = style["bold"].toBool();


    if(width < 100) width = 100;


    doc.setColumnWidth(column, width);
    doc.setRowHeight(row, height);
    QXlsx::Format format;
    QFont font(fontFamily);
    font.setBold(bold);
    font.setPointSize(fontSize);
    format.setFont(font);
    format.setFontColor(QColor(color));
    format.setTextWrap(true);
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

void ExcelJsonTable::writeRow(QJsonArray Row)
{
    QJsonObject obj;
    int columnCount = getMaxColumnCount(sheetIndex);
    int rowCount = Row.size();
    bool columnToRow = ( rowCount < columnCount)? true : false;
    for(int i=0; i < Row.size(); i++)
    {
        obj = Row[i].toObject();
        if(obj.value("type").toString().compare("img", Qt::CaseInsensitive) == 0)
            if(skipImages)
                continue;

        if(columnToRow)
        {
            QXlsx::CellRange range(currentRow, 1, currentRow, columnCount);
            doc.mergeCells(range);
            writeCell(currentRow, currentColumn, obj);
            currentRow++;
        }
        else
        {
            writeCell(currentRow, currentColumn, obj);
            currentColumn++;
        }
    }
    if(!columnToRow)
        currentRow++;
    currentColumn = 1;

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
    if(sheetIndex < 1) sheetIndex = 1;
    // find sheetIndex empty row
    if(sheetIndex == 1) return 1;

    int index = 1, temp, row = 0;
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

