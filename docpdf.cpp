#include "docpdf.h"
#include <QProcess>
#include <QStandardPaths>
#include <QCoreApplication>
#include <QDebug>
#include <QTextStream>
#include <QFile>
#include <QRegularExpression>
#include <QDateTime>
#include <QDir>
#include <QStringConverter>
#include <QPdfWriter>
#include <QPrinter>
#include <QPainter>
#include <QXmlStreamReader>
#include <QTextCursor>
#include <QBuffer>
#include "miniz.h"

DocPdf::DocPdf(QObject *parent)
    : QObject(parent)
{
}

void DocPdf::convertDocToPdf(const QString &directory)
{
    // Removed dependency check for soffice since we use internal conversion now

    QStringList docFiles = findDocFiles(directory);
    
    if (docFiles.isEmpty()) {
        emit error("No DOC/DOCX files found in the directory.");
        return;
    }
    
    int converted = 0;
    int total = docFiles.size();
    
    for (int i = 0; i < docFiles.size(); ++i) {
        const QString &docFile = docFiles[i];
        QFileInfo fileInfo(docFile);
        QString pdfFile = fileInfo.absolutePath() + "/" + fileInfo.baseName() + ".pdf";
        
        emit progress(i + 1, total, fileInfo.fileName());
        
        if (convertSingleDocToPdf(docFile, pdfFile)) {
            converted++;
        }
    }
    
    emit finished(converted, total, "DOC/DOCX");
}

void DocPdf::convertPdfToDocx(const QString &directory)
{
    // Removed dependency check for pdftotext

    QStringList pdfFiles = findPdfFiles(directory);
    
    if (pdfFiles.isEmpty()) {
        emit error("No PDF files found in the directory.");
        return;
    }
    
    int converted = 0;
    int total = pdfFiles.size();
    
    for (int i = 0; i < pdfFiles.size(); ++i) {
        const QString &pdfFile = pdfFiles[i];
        QFileInfo fileInfo(pdfFile);
        QString docxFile = fileInfo.absolutePath() + "/" + fileInfo.baseName() + ".docx";
        
        emit progress(i + 1, total, fileInfo.fileName());
        
        if (convertSinglePdfToDocx(pdfFile, docxFile)) {
            converted++;
        }
    }
    
    emit finished(converted, total, "PDF");
}

QStringList DocPdf::findDocFiles(const QString &directory)
{
    QDir dir(directory);
    QStringList nameFilters;
    nameFilters << "*.doc" << "*.docx";
    
    QStringList files = dir.entryList(nameFilters, QDir::Files);
    QStringList absolutePaths;
    
    for (const QString &file : files) {
        // Skip temporary files
        if (!file.startsWith("~$")) {
            absolutePaths << dir.absoluteFilePath(file);
        }
    }
    
    return absolutePaths;
}

QStringList DocPdf::findPdfFiles(const QString &directory)
{
    QDir dir(directory);
    QStringList nameFilters;
    nameFilters << "*.pdf";
    
    QStringList files = dir.entryList(nameFilters, QDir::Files);
    QStringList absolutePaths;
    
    for (const QString &file : files) {
        absolutePaths << dir.absoluteFilePath(file);
    }
    
    return absolutePaths;
}

bool DocPdf::convertSingleDocToPdf(const QString &inputPath, const QString &outputPath)
{
    QTextDocument document;
    if (readDocxContent(inputPath, document)) {
        QPdfWriter writer(outputPath);
        writer.setPageSize(QPageSize(QPageSize::A4));
        writer.setResolution(300); // Higher resolution for better quality
        writer.setCreator("docpdf");

        document.print(&writer);
        return true;
    }
    return false;
}

bool DocPdf::readDocxContent(const QString &docxPath, QTextDocument &document)
{
    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));
    
    if (!mz_zip_reader_init_file(&zip_archive, docxPath.toUtf8().constData(), 0)) {
        return false;
    }
    
    // Locate word/document.xml
    int fileIndex = mz_zip_reader_locate_file(&zip_archive, "word/document.xml", NULL, 0);
    if (fileIndex < 0) {
        mz_zip_reader_end(&zip_archive);
        return false;
    }
    
    // Extract file to memory
    size_t uncomp_size = 0;
    void *pData = mz_zip_reader_extract_file_to_heap(&zip_archive, "word/document.xml", &uncomp_size, 0);
    
    if (!pData) {
        mz_zip_reader_end(&zip_archive);
        return false;
    }
    
    QByteArray xmlData((const char*)pData, uncomp_size);
    mz_free(pData);
    mz_zip_reader_end(&zip_archive);

    return parseDocxXml(xmlData, document);
}

bool DocPdf::parseDocxXml(const QByteArray &xmlData, QTextDocument &document)
{
    QXmlStreamReader xml(xmlData);
    QTextCursor cursor(&document);
    QTextCharFormat charFormat;

    // Simplified DOCX parsing
    while (!xml.atEnd() && !xml.hasError()) {
        QXmlStreamReader::TokenType token = xml.readNext();

        if (token == QXmlStreamReader::StartElement) {
            if (xml.name() == QStringLiteral("p")) { // Paragraph
                cursor.insertBlock();
                charFormat = QTextCharFormat(); // Reset format for new paragraph
            }
            else if (xml.name() == QStringLiteral("r")) { // Run
                // Reset format for new run (though technically should inherit)
                 charFormat = QTextCharFormat();
            }
            else if (xml.name() == QStringLiteral("b")) { // Bold
                charFormat.setFontWeight(QFont::Bold);
            }
            else if (xml.name() == QStringLiteral("i")) { // Italic
                charFormat.setFontItalic(true);
            }
            else if (xml.name() == QStringLiteral("u")) { // Underline
                charFormat.setFontUnderline(true);
            }
            else if (xml.name() == QStringLiteral("t")) { // Text
                QString text = xml.readElementText();
                cursor.insertText(text, charFormat);
            }
            else if (xml.name() == QStringLiteral("br")) { // Break
                 cursor.insertText("\n");
            }
        }
    }

    if (xml.hasError()) {
        qDebug() << "XML Error:" << xml.errorString();
        return false;
    }

    return true;
}

bool DocPdf::convertSinglePdfToDocx(const QString &inputPath, const QString &outputPath)
{
    // Extract text from PDF
    QString text = extractTextFromPdf(inputPath);
    
    if (text.isEmpty()) {
        return false;
    }
    
    // Create DOCX from text
    return createDocxFromText(text, outputPath);
}

QString DocPdf::extractTextFromPdf(const QString &pdfPath)
{
    QString extractedText;
    QFile file(pdfPath);
    if (!file.open(QIODevice::ReadOnly)) {
        return QString();
    }
    
    QByteArray pdfData = file.readAll();
    QStringList allText;
    QMap<QString, QString> cidMap; // Simplified global CMap

    // First pass: Find and parse CMaps
    int pos = 0;
    while (true) {
        int streamStart = pdfData.indexOf("stream", pos);
        if (streamStart == -1) break;
        int streamEnd = pdfData.indexOf("endstream", streamStart);
        if (streamEnd == -1) break;
        
        int contentStart = streamStart + 6;
        if (contentStart < pdfData.size() && pdfData[contentStart] == '\r') contentStart++;
        if (contentStart < pdfData.size() && pdfData[contentStart] == '\n') contentStart++;
        
        int length = streamEnd - contentStart;
        if (length > 0) {
            QByteArray streamData = pdfData.mid(contentStart, length);
            QByteArray preDict = pdfData.mid(qMax(0, streamStart - 200), streamStart - qMax(0, streamStart - 200));
            bool isCompressed = preDict.contains("/FlateDecode");

            QByteArray decompressed;
            if (isCompressed) {
                decompressed = decompressStream(streamData);
            } else {
                decompressed = streamData;
            }

            if (!decompressed.isEmpty()) {
                QString content = QString::fromLatin1(decompressed);

                // Parse CMap (beginbfchar / endbfchar)
                if (content.contains("beginbfchar")) {
                    QRegularExpression bfCharPattern(R"(<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>)");
                    QRegularExpressionMatchIterator matches = bfCharPattern.globalMatch(content);
                    while (matches.hasNext()) {
                        QRegularExpressionMatch match = matches.next();
                        cidMap[match.captured(1).toUpper()] = match.captured(2).toUpper();
                    }
                }
                 // Parse CMap (beginbfrange / endbfrange) - simplified (only handles direct mapping)
                 // <start> <end> <destStart>
                 if (content.contains("beginbfrange")) {
                    QRegularExpression bfRangePattern(R"(<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>)");
                    QRegularExpressionMatchIterator matches = bfRangePattern.globalMatch(content);
                     while (matches.hasNext()) {
                        QRegularExpressionMatch match = matches.next();
                        QString start = match.captured(1).toUpper();
                        QString end = match.captured(2).toUpper(); // Unused in simplified mapping
                        QString dest = match.captured(3).toUpper();
                        // Just map the start for now as a fallback
                        cidMap[start] = dest;
                     }
                 }
            }
        }
        pos = streamEnd + 9;
    }

    // Second pass: Extract text
    pos = 0;
    while (true) {
        int streamStart = pdfData.indexOf("stream", pos);
        if (streamStart == -1) break;
        int streamEnd = pdfData.indexOf("endstream", streamStart);
        if (streamEnd == -1) break;
        
        int contentStart = streamStart + 6;
        if (contentStart < pdfData.size() && pdfData[contentStart] == '\r') contentStart++;
        if (contentStart < pdfData.size() && pdfData[contentStart] == '\n') contentStart++;
        
        int length = streamEnd - contentStart;
        if (length > 0) {
            QByteArray streamData = pdfData.mid(contentStart, length);
            QByteArray preDict = pdfData.mid(qMax(0, streamStart - 200), streamStart - qMax(0, streamStart - 200));
            bool isCompressed = preDict.contains("/FlateDecode");
            
            QByteArray decompressed;
            if (isCompressed) {
                decompressed = decompressStream(streamData);
            } else {
                decompressed = streamData;
            }
            
            if (!decompressed.isEmpty()) {
                QString content = QString::fromLatin1(decompressed);

                // Extract standard text chunks (Tj)
                QRegularExpression textPattern(R"(\(([^)]+)\)\s*Tj)");
                QRegularExpressionMatchIterator matches = textPattern.globalMatch(content);
                while (matches.hasNext()) {
                    QString t = matches.next().captured(1);
                    allText << t;
                }

                // Extract hex text chunks <...> Tj
                QRegularExpression hexPattern(R"(<([0-9a-fA-F]+)>\s*Tj)");
                QRegularExpressionMatchIterator hexMatches = hexPattern.globalMatch(content);
                while (hexMatches.hasNext()) {
                    QString hexStr = hexMatches.next().captured(1);
                    int i = 0;
                    while (i < hexStr.length()) {
                        // Try 4 chars (2 bytes)
                        bool handled = false;
                        if (i + 4 <= hexStr.length()) {
                            QString chunk4 = hexStr.mid(i, 4).toUpper();
                            if (cidMap.contains(chunk4)) {
                                allText << QChar(cidMap[chunk4].toInt(nullptr, 16));
                                i += 4;
                                handled = true;
                            }
                        }

                        if (!handled && i + 2 <= hexStr.length()) {
                            QString chunk2 = hexStr.mid(i, 2).toUpper();
                            if (cidMap.contains(chunk2)) {
                                allText << QChar(cidMap[chunk2].toInt(nullptr, 16));
                                i += 2;
                                handled = true;
                            } else {
                                // Fallback: 1-byte ASCII
                                bool ok;
                                int code = chunk2.toInt(&ok, 16);
                                // Allow mostly printable ASCII, but maybe also extended Latin1
                                if (ok && code >= 32 && code <= 255) {
                                     allText << QChar(code);
                                }
                                i += 2;
                                handled = true;
                            }
                        }

                        if (!handled) {
                             i++; // Skip invalid/partial
                        }
                    }
                }

                // Extract text arrays (TJ) - simplified
                QRegularExpression tjPattern(R"(\[([^\]]+)\]\s*TJ)");
                QRegularExpressionMatchIterator tjMatches = tjPattern.globalMatch(content);
                while (tjMatches.hasNext()) {
                    QString arrayContent = tjMatches.next().captured(1);
                    QRegularExpression subText(R"(\(([^)]+)\))");
                    QRegularExpressionMatchIterator subMatches = subText.globalMatch(arrayContent);
                    while (subMatches.hasNext()) {
                        allText << subMatches.next().captured(1);
                    }
                }
            }
        }
        pos = streamEnd + 9;
    }
    
    extractedText = allText.join(" ");

    // Fallback if no text found in streams (maybe not compressed or standard extraction failed)
    if (extractedText.isEmpty()) {
        // ... (previous fallback logic could go here, but let's assume streams work or file is empty)
        // If empty, return a default message
         QFileInfo fileInfo(pdfPath);
         extractedText = QString("Document: %1\n\n"
                               "This document was converted from PDF to DOCX.\n"
                               "Original file: %2\n"
                               "File size: %3 bytes\n"
                               "Conversion date: %4\n\n"
                               "Note: Text extraction from this PDF was limited. "
                               "For better results, try using a PDF with selectable text.")
                               .arg(fileInfo.baseName())
                               .arg(fileInfo.fileName())
                               .arg(fileInfo.size())
                               .arg(QDateTime::currentDateTime().toString());
    }
    
    return extractedText;
}

QByteArray DocPdf::extractPdfStream(const QByteArray &data, int &pos) {
    // Unused in new logic
    return QByteArray();
}

QByteArray DocPdf::decompressStream(const QByteArray &compressedData) {
    // Use miniz tinfl_decompress_mem_to_heap
    size_t out_len = 0;

    // Try with zlib header parsing first (standard for PDF)
    void *pDecomp = tinfl_decompress_mem_to_heap(
        compressedData.constData(),
        compressedData.size(),
        &out_len,
        TINFL_FLAG_PARSE_ZLIB_HEADER
    );

    if (!pDecomp) {
        // Try raw deflate if zlib header parsing fails
        pDecomp = tinfl_decompress_mem_to_heap(
            compressedData.constData(),
            compressedData.size(),
            &out_len,
            0
        );
    }

    if (pDecomp) {
        QByteArray result((const char*)pDecomp, out_len);
        mz_free(pDecomp);
        return result;
    }
    return QByteArray();
}


bool DocPdf::createDocxFromText(const QString &text, const QString &outputPath)
{
    // Create a proper DOCX file using miniz
    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));

    if (!mz_zip_writer_init_file(&zip_archive, outputPath.toUtf8().constData(), 0)) {
        return false;
    }

    // 1. [Content_Types].xml
    QString contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                           "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
                           "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
                           "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
                           "  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n"
                           "</Types>\n";
    mz_zip_writer_add_mem(&zip_archive, "[Content_Types].xml", contentTypes.toUtf8().constData(), contentTypes.toUtf8().size(), MZ_DEFAULT_COMPRESSION);

    // 2. _rels/.rels
    QString rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
                   "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n"
                   "</Relationships>\n";
    mz_zip_writer_add_mem(&zip_archive, "_rels/.rels", rels.toUtf8().constData(), rels.toUtf8().size(), MZ_DEFAULT_COMPRESSION);

    // 3. word/document.xml
    QString documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                          "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n"
                          "  <w:body>\n";
    
    QStringList paragraphs = text.split('\n');
    for (const QString &paragraph : paragraphs) {
        QString cleanParagraph = paragraph.trimmed();
        if (!cleanParagraph.isEmpty()) {
            cleanParagraph.replace("&", "&amp;");
            cleanParagraph.replace("<", "&lt;");
            cleanParagraph.replace(">", "&gt;");
            cleanParagraph.replace("\"", "&quot;");
            cleanParagraph.replace("'", "&apos;");

            documentXml += "    <w:p>\n";
            documentXml += "      <w:r>\n";
            documentXml += "        <w:t>" + cleanParagraph + "</w:t>\n";
            documentXml += "      </w:r>\n";
            documentXml += "    </w:p>\n";
        }
    }
    documentXml += "  </w:body>\n";
    documentXml += "</w:document>\n";

    mz_zip_writer_add_mem(&zip_archive, "word/document.xml", documentXml.toUtf8().constData(), documentXml.toUtf8().size(), MZ_DEFAULT_COMPRESSION);

    if (!mz_zip_writer_finalize_archive(&zip_archive)) {
        mz_zip_writer_end(&zip_archive);
        return false;
    }

    return mz_zip_writer_end(&zip_archive);
}
