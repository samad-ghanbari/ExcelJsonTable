#ifndef TABLETEMPLATE_H
#define TABLETEMPLATE_H

#include <QJsonObject>
#include <QJsonArray>
#include "jsontable.h"
#include <QStringList>
#include <QByteArray>

class TableTemplate
{
public:
    explicit TableTemplate(int _columnCount, QJsonObject _headerStyle={}, QJsonObject _contentStyle={} ) :
        table("#000", "#FFF","tahoma",14), title("#000", "#DDD","Tahoma",14)
    {
        columnCount = _columnCount;

        if(_headerStyle.isEmpty())
        {
            headerStyle = table.createStyle("Table Header",0,0,"#558","#EEF","Tahoma",14,true,"center",1,0,2);
        }
        else
            headerStyle = _headerStyle;

        if(_contentStyle.isEmpty())
        {
            contentStyle = table.createStyle("Table Content",0,0,NULL,NULL,NULL,0,false,"left");
        }
        else
            contentStyle = _contentStyle;
    }

    bool appendTitle(QStringList types, QStringList values, QStringList colors, QStringList aligns)
    {
        if(types.size() != values.size())
            return false;
        if(types.size() != colors.size())
            return false;
        if(aligns.size() != colors.size())
            return false;

        QJsonArray row;
        QJsonObject style;
        for(int i = 0; i < types.size(); i++)
        {
            if(types[i].compare("img", Qt::CaseInsensitive) == 0)
                style = title.createStyle("Header Image", 0, 0, NULL, NULL, NULL,0,false,aligns[i], 0);
            else
                style = title.createStyle("header Text", 0, 0, colors[i], NULL, "Tahoma", 16, true, aligns[i], 0 );

            QJsonObject obj = title.createObject(types[i], values[i], style);
            row = title.addObjectToRow(row, obj);
        }

        title.addRowToTable(row);
        return true;
    }

    bool appendTitle(QStringList types, QStringList values, QList<double> widths, QList<double> heights, QStringList colors, QStringList aligns)
    {
        if(types.size() != values.size())
            return false;
        if(types.size() != colors.size())
            return false;
        if(aligns.size() != colors.size())
            return false;
        if(colors.size() != widths.size())
            return false;
        if(widths.size() != heights.size())
            return false;

        QJsonArray row;
        QJsonObject style;
        for(int i = 0; i < types.size(); i++)
        {
            if(types[i].compare("img", Qt::CaseInsensitive) == 0)
                style = title.createStyle("Header Image", widths[i], heights[i], NULL, NULL, NULL,0,false,aligns[i], 0);
            else
                style = title.createStyle("header Text", widths[i], heights[i], colors[i], NULL, "Tahoma", 16, true, aligns[i], 0 );

            QJsonObject obj = title.createObject(types[i], values[i], style);
            row = title.addObjectToRow(row, obj);
        }

        title.addRowToTable(row);
        return true;
    }

    QJsonArray getTitle(double width)
    {
        title.updateSameWidth(width);
        title.updateHeight();
        title.updateRowSpan(false);

        return title.table;
    }

    QByteArray getTitleByteArray()
    {
        return title.toByteArray();
    }

    bool setTableHeader(QStringList headerList)
    {

        if(headerList.size() == columnCount)
        {
            QJsonArray row = table.createObjects("text", headerList, headerStyle);
            table.addRowToTable(row);
        }
        else
            return false;

        return true;
    }

    bool appentRow(QStringList rowList)
    {
        if(rowList.size() == columnCount)
        {
            QJsonArray row = table.createObjects("text", rowList,contentStyle);
            table.addRowToTable(row);
        }
        else
            return false;

        return true;
    }

    bool appendRow(QStringList types, QStringList values, QList<double> widths, QList<double> heights, QStringList colors, QStringList aligns)
    {
        if(types.size() != values.size())
            return false;
        if(types.size() != colors.size())
            return false;
        if(aligns.size() != colors.size())
            return false;
        if(colors.size() != widths.size())
            return false;
        if(widths.size() != heights.size())
            return false;

        QJsonArray row;
        QJsonObject style;
        for(int i = 0; i < types.size(); i++)
        {
            if(types[i].compare("img", Qt::CaseInsensitive) == 0)
                style = table.createStyle("Table Image", widths[i], heights[i], NULL, NULL, NULL,0,false,aligns[i], 0);
            else
                style = table.createStyle("Table Content", widths[i],heights[i], colors[i], NULL, "Tahoma", 16, true, aligns[i], 0 );

            QJsonObject obj = table.createObject(types[i], values[i], style);
            row = table.addObjectToRow(row, obj);
        }

        table.addRowToTable(row);
        return true;
    }

    void highlight(int row)
    {
        table.highlightRow(row);
    }

    QJsonArray getTable(double width, int columnSpanNumber, bool justifyWidth = true)
    {
        if(justifyWidth)
        {
            table.updateFairCell(width, true);
            table.updateRowSpan(columnSpanNumber);
        }
        else
        {
            table.updateSameWidth(width);
            table.updateHeight();
            table.updateRowSpan(true,columnSpanNumber);
        }


        return table.table;
    }

    QByteArray getTableByteArray()
    {
        return table.toByteArray();
    }

    int columnCount;
    QJsonObject headerStyle, contentStyle;
    JsonTable table, title;
};

#endif // TABLETEMPLATE_H
