#include <QCoreApplication>
#include <QXlsx/header/xlsxdocument.h>
#include <QXlsx/header/xlsxchartsheet.h>
#include <QXlsx/header/xlsxcellrange.h>
#include <QXlsx/header/xlsxchart.h>
#include <QXlsx/header/xlsxrichstring.h>
#include <QXlsx/header/xlsxworkbook.h>

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    QXlsx::Document xlsx;
    xlsx.write(1,1,QVariant("Hello World"));
    xlsx.saveAs("file.xlsx");

    return 1;
}
