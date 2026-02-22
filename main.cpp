#include <QApplication>
#include <QTimer>
#include <QDir>
#include <QPdfWriter>
#include <QPainter>
#include <QDebug>
#include "mainwindow.h"
#include "docpdf.h"

void runTest() {
    qDebug() << "Running self-test...";

    QString testDir = QDir::currentPath() + "/test_data";
    QDir().mkpath(testDir);

    // 1. Generate a test PDF
    QString pdfPath = testDir + "/test_input.pdf";
    QPdfWriter writer(pdfPath);
    writer.setPageSize(QPageSize(QPageSize::A4));
    QPainter painter(&writer);
    // Use a font that hopefully doesn't use complex encoding/subsetting if possible
    QFont font("Times New Roman", 12);
    painter.setFont(font);
    painter.drawText(100, 100, "Hello World from PDF!");
    painter.drawText(100, 300, "(This text is in parentheses)");
    painter.end();

    qDebug() << "Created test PDF at" << pdfPath;

    // 2. Convert PDF to DOCX
    DocPdf converter;
    qDebug() << "Converting PDF to DOCX...";
    converter.convertPdfToDocx(testDir);

    QString docxPath = testDir + "/test_input.docx";
    if (QFile::exists(docxPath)) {
        qDebug() << "SUCCESS: DOCX created at" << docxPath;
    } else {
        qCritical() << "FAILURE: DOCX not created";
        QCoreApplication::exit(1);
        return;
    }

    // 3. Convert DOCX back to PDF
    qDebug() << "Converting DOCX to PDF...";
    // Rename original PDF to avoid confusion/overwriting check if needed,
    // but convertDocToPdf will write test_input.pdf, checking if it overwrites or fails.
    // Let's rename the original PDF to verify the new one is created.
    QFile::rename(pdfPath, testDir + "/original.pdf");

    converter.convertDocToPdf(testDir);

    if (QFile::exists(pdfPath)) {
        qDebug() << "SUCCESS: PDF created at" << pdfPath;
    } else {
        qCritical() << "FAILURE: PDF not created";
        QCoreApplication::exit(1);
        return;
    }

    qDebug() << "All tests passed.";
    QCoreApplication::exit(0);
}

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    
    app.setApplicationName("docpdf");
    app.setApplicationVersion("1.0.0");
    app.setOrganizationName("DocPDF");
    
    if (app.arguments().contains("--test")) {
        QTimer::singleShot(0, runTest);
        return app.exec();
    }

    MainWindow window;
    window.show();
    
    return app.exec();
}