#include "xlsxstyles_p.h"
#include <QCoreApplication>
#include <QXlsx/header/xlsxdocument.h>
#include <QXlsx/header/xlsxchartsheet.h>
#include <QXlsx/header/xlsxcellrange.h>
#include <QXlsx/header/xlsxchart.h>
#include <QXlsx/header/xlsxrichstring.h>
#include <QXlsx/header/xlsxworkbook.h>

#include <QDebug>

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    QDate today(QDate::currentDate());
    QXlsx::Document xlsx;
    xlsx.addSheet(today.toString("yyyy-MM-dd"));
    xlsx.deleteSheet("Sheet1");

    //xlsx.currentWorksheet()->setGridLinesVisible(false);
    xlsx.write(2,2,QVariant("New Sheet"));
    xlsx.setRowHeight(1,40);
    xlsx.setRowHeight(2,40);
    xlsx.setColumnWidth(1, 30);
    xlsx.setColumnWidth(2, 40);
    QXlsx::Format fmt;
    fmt.setFontBold(true);
    QFont font("times");
    //font.setPixelSize(18);
    //fmt.setFontSize(18);
    font.setPointSize(18);
    fmt.setFont(font);
    fmt.setFontColor(QColor(Qt::red));
    fmt.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    fmt.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    fmt.setBottomBorderColor(Qt::blue);
    fmt.setFillPattern(QXlsx::Format::PatternSolid);
    fmt.setPatternBackgroundColor(Qt::lightGray);
    //fmt.setTopBorderStyle(QXlsx::Format::BorderThick);
    fmt.setBorderStyle(QXlsx::Format::BorderThick);

    xlsx.mergeCells(QXlsx::CellRange(1, 1, 1, 4));

    //xlsx.setColumnFormat(1, fmt);
    xlsx.write(1,1,QVariant("Hello World"),fmt);

    xlsx.addSheet(today.toString("yyyy"));
    xlsx.write(1,1,QVariant("Hello World 2"),fmt);

    QVariant val = xlsx.read(1, 1);
    qDebug() << val.toString();

    xlsx.selectSheet(0);
    val = xlsx.read(1, 1);
    qDebug() << val.toString();

    xlsx.saveAs("file.xlsx");



    // read

    QXlsx::Document doc("a.xlsx");

        if(doc.load())
    {
        QXlsx::CellRange CR = doc.dimension();
        int firstColumn = CR.firstColumn();
        int lastColumn = CR.lastColumn();
        int firstRow = CR.firstRow();
        int lastRow = CR.lastRow();
        qDebug() << firstRow << lastRow;
        qDebug() << firstColumn << lastColumn;
        qDebug()<< "row count" << CR.rowCount();

        lastRow = (lastRow > 20)? 20: lastRow;
        qDebug() << doc.read(3,3).toString();

        for(int r=firstRow; r<=lastRow; r++)
        {
            for(int c=firstColumn; c<=lastColumn; c++)
            {
                val = doc.read(r, c);
                qDebug() << val.toString();
            }
        }
    }

    return 1;
}
