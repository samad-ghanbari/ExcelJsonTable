#include <QApplication>
#include "exceljsontable.h"
#include "lib/jsontable.h"
#include <QJsonArray>
#include <QJsonObject>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    JsonTable title(40,"#00A","#eee","tahoma", 16);
    JsonTable table(40,"#00A","#eee","tahoma", 16);

    //title repeated on eatch sheet
    QJsonArray Row;
    QJsonObject style = title.createStyle("DSLAM", 0, 40, "#00A", "#eee","tahoma", 18,true,"center",1);
    QJsonObject obj = title.createObject("text", "Area 2 DSLAM Plan", style);
    Row = title.addObjectToRow(Row, obj);
    title.addRowToTable(Row);

    //content table
    style = table.createStyle("DSLAM Connections", 0, 20, "#008", "#eef","tahoma", 14,true,"center",1);
    Row = table.createObjects("text", {"Exchange", "Saloon", "Interface", "Module"}, style);
    table.addRowToTable(Row);

    style = table.createStyle("DSLAM Connections", 0, 20, "#080", "#efe","tahoma", 14,false,"left",1);
    Row = table.createObjects("text", {"BA1", "Switch1", "1G 0/7/1", "SX1 500m"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA2", "Switch2", "1G 0/7/2", "SX2 500m"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA3", "Switch3", "1G 0/7/3", "SX3 500m"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA4", "Switch4", "1G 0/7/4", "SX4 500m"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA5", "Switch5", "1G 0/7/5", "SX5 500m"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA6", "Switch6", "1G 0/7/6", "SX6 500m"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA7", "Switch7", "1G 0/7/7", "SX7 500m"}, style);
    table.addRowToTable(Row);

    table.addRowToTable(); // new sheet

    style = table.createStyle("Metro Connections", 0, 20, "#808", "#fef","tahoma", 14,true,"center",1);
    Row = table.createObjects("text", {"Exchange", "Device", "Interface"}, style);
    table.addRowToTable(Row);


    style = table.createStyle("Metro Connections", 0, 20, "#080", "#ffe","tahoma", 14,false,"left",1);
    Row = table.createObjects("text", {"BA1", "CX600-x16", "1G 1/0/0"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA2", "CX600-x8", "1G 2/0/0"}, style);
    table.addRowToTable(Row);

    Row = table.createObjects("text", {"BA3", "S12700", "1G 3/0/0"}, style);
    table.addRowToTable(Row);


    table.updateTableRowHeight();

    ExcelJsonTable ejs(title.table, table.table);

    ejs.exportExcel("test.xlsx");


    return 1;
}
