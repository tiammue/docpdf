#ifndef DOCPDF_H
#define DOCPDF_H

#include <QObject>
#include <QString>
#include <QStringList>
#include <QDir>
#include <QFileInfo>
#include <QTextDocument>

class DocPdf : public QObject
{
    Q_OBJECT

public:
    explicit DocPdf(QObject *parent = nullptr);

public slots:
    void convertDocToPdf(const QString &directory);
    void convertPdfToDocx(const QString &directory);

signals:
    void progress(int current, int total, const QString &filename);
    void finished(int converted, int total, const QString &type);
    void error(const QString &errorMessage);

private:
    QStringList findDocFiles(const QString &directory);
    QStringList findPdfFiles(const QString &directory);

    // Internal conversion logic
    bool convertSingleDocToPdf(const QString &inputPath, const QString &outputPath);
    bool convertSinglePdfToDocx(const QString &inputPath, const QString &outputPath);

    // DOCX -> PDF Helpers
    bool readDocxContent(const QString &docxPath, QTextDocument &document);
    bool parseDocxXml(const QByteArray &xmlData, QTextDocument &document);

    // PDF -> DOCX Helpers
    QString extractTextFromPdf(const QString &pdfPath);
    bool createDocxFromText(const QString &text, const QString &outputPath);
    QByteArray extractPdfStream(const QByteArray &data, int &pos);
    QByteArray decompressStream(const QByteArray &compressedData);
};

#endif // DOCPDF_H