#ifndef DOCPDF_H
#define DOCPDF_H

#include <QObject>
#include <QString>
#include <QStringList>
#include <QDir>
#include <QFileInfo>

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
    bool convertSingleDocToPdf(const QString &inputPath, const QString &outputPath);
    bool convertSinglePdfToDocx(const QString &inputPath, const QString &outputPath);
    QString extractTextFromPdf(const QString &pdfPath);
    bool createDocxFromText(const QString &text, const QString &outputPath);
};

#endif // DOCPDF_H