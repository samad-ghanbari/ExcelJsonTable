#include <QApplication>
#include "exceljsontable.h"
#include <QJsonArray>
#include <QJsonObject>
#include <lib/tableTemplate.h>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    TableTemplate tableTemplate(4);
    // page title
    tableTemplate.appendTitle({"img", "text","img"}, {":/danet.png", "DaNet Report", ":/danet.png"},{"", "#005", ""}, {"left", "center", "right"});
    tableTemplate.appendTitle({"text", "text"}, {"2-BA Saloon Data", "1403/05/29"}, {"#050", "#500"}, {"left", "right"});

    //table
    tableTemplate.setTableHeader({"Menu-1", "Menu-2", "Menu-3", "Menu-4"});

    tableTemplate.appentRow({"Lorem-1 ipsum odor amet, consectetuer adipiscing elit. Blandit eget vulputate cubilia convallis penatibus vivamus ante ante? Odio felis libero auctor elit, parturient donec porta tristique nullam. Scelerisque penatibus maximus erat aptent egestas mus. Eu sed euismod, hac semper arcu tortor ullamcorper vestibulum.", "Lorem-1", "Ipsum-1", "L-I-1"});
    tableTemplate.appentRow({"Lorem-2", "Data", "CX600X16", "1G 1/2/3"});

    tableTemplate.appentRow({"Lorem-3 ipsum", "PCM", "CX600X16", "Lorem ipsum odor amet, consectetuer adipiscing elit. Non arcu scelerisque nascetur elementum ante iaculis sapien. Facilisi faucibus dolor arcu ante consequat accumsan facilisi."});
    tableTemplate.appentRow({"lorem-4", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-5 ", "Data", "CX600X16", "1G 1/1/1"});

    tableTemplate.appendRow({"text", "img", "text", "text"}, {"samad",":/tct.png", "test2", "test3"},{0,0,0,0}, {0,0,0,0},{"#F00", "#F55", "#F33", "#F80"},{"center", "right", "center", "right"});

    tableTemplate.appentRow( {"Lorem-6", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-7", "Switch", "Lorem ipsum odor amet, consectetuer adipiscing elit.Lorem ipsum odor amet, consectetuer adipiscing elit.Lorem ipsum odor amet, consectetuer adipiscing elit.", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-8", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-9", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-10", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-11", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-12", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-13", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-14", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-15", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-16", "Switch", "CX600X16 NetEngine Huawei", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-17", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-18", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-1 ipsum odor amet, consectetuer adipiscing elit. Blandit eget vulputate cubilia convallis penatibus vivamus ante ante? Odio felis libero auctor elit, parturient donec porta tristique nullam. Scelerisque penatibus maximus erat aptent egestas mus. Eu sed euismod, hac semper arcu tortor ullamcorper vestibulum.", "Lorem-1", "Ipsum-1", "L-I-1"});
    tableTemplate.appentRow({"Lorem-3 ipsum", "PCM", "CX600X16", "Lorem ipsum odor amet, consectetuer adipiscing elit. Non arcu scelerisque nascetur elementum ante iaculis sapien. Facilisi faucibus dolor arcu ante consequat accumsan facilisi."});
    tableTemplate.appentRow({"lorem-4", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow( {"Lorem-6", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-7", "Switch", "Lorem ipsum odor amet, consectetuer adipiscing elit.Lorem ipsum odor amet, consectetuer adipiscing elit.Lorem ipsum odor amet, consectetuer adipiscing elit.", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-8", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-9", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-10", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-11", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-12", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-13", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-14", "PCM", "CX600X16", "10G 3/0/0"});
    tableTemplate.appentRow({"Lorem-15", "Switch", "CX600X16", "1G 10/0/0"});
    tableTemplate.appentRow({"Lorem-16", "Switch", "CX600X16 NetEngine Huawei", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-17", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appentRow({"Lorem-18", "Switch", "CX600X16", "10G 11/0/0"});
    tableTemplate.appendRow({"text", "img", "text", "text"}, {"samad",":/tct.png", "test2", "test3"},{0,0,0,0}, {0,0,0,0},{"#F00", "#F55", "#F33", "#F80"},{"center", "right", "center", "right"});


    tableTemplate.highlight(5);

    QJsonArray table = tableTemplate.getTable(1500, 2);
    QJsonArray title = tableTemplate.getTitle(1500);

    ExcelJsonTable ejs;
    ejs.setTables(title,table);
    ejs.exportExcel("file.xlsx",{0},false);

    //QJsonArray jt = ejs.excelToJson("a.xlsx",0,3,2);

    return 1;
}
