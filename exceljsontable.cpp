#include "exceljsontable.h"

ExcelJsonTable::ExcelJsonTable(QJsonArray _sheetTitle, QJsonArray _table, QObject *parent)
    : QObject{parent}
{
    tableArray = _table;
    sheetTitle = _sheetTitle;

}

void ExcelJsonTable::exportExcel(QString _outputPath)
{
    outputPath = _outputPath;
    int sheetCount = getSheetCount();
    if(sheetCount == 0) return;
    int start, row=1;
    QJsonArray Row;
    QJsonObject obj;


    doc.saveAs(outputPath);
}

void ExcelJsonTable::writeCell(int sheetIndex, int row, int column, QJsonObject Obj)
{
    doc.selectSheet(sheetIndex);
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

bool ExcelJsonTable::addSheet(QString sheetName)
{
    if(doc.addSheet(sheetName))
        return true;
    else
        return false;
}

bool ExcelJsonTable::removeSheet(QString sheetName)
{
    if(doc.deleteSheet(sheetName))
        return true;
    else
        return false;
}

QString ExcelJsonTable::getSheetName(int sheetIndex)
{
    int row = getStartRow(sheetIndex);
    QJsonArray Row = tableArray[row].toArray();
    QString name = QString("sheet-") + QString::number(sheetIndex);
    if(Row.count() > 0)
    {
        QJsonObject obj = Row[0].toObject()["style"].toObject();
        name = obj["name"].toString();
    }
    return name;
}

int ExcelJsonTable::getStartRow(int sheetIndex)
{
    if(sheetIndex < 0) sheetIndex =0;

    int index = -1, temp = 0, row = 0;
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
    int temp, count = 0;
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

