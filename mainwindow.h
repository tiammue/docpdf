#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QPushButton>
#include <QLabel>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QProgressBar>
#include <QTextEdit>
#include <QThread>
#include "docpdf.h"

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void convertDocToPdf();
    void convertPdfToDocx();
    void onConversionProgress(int current, int total, const QString &filename);
    void onConversionFinished(int converted, int total, const QString &type);
    void onConversionError(const QString &error);

private:
    void setupUI();
    void updateStatus(const QString &message, const QString &color = "black");
    QString getCurrentDirectory();
    
    // UI Components
    QWidget *m_centralWidget;
    QVBoxLayout *m_mainLayout;
    QLabel *m_titleLabel;
    QLabel *m_dirLabel;
    QHBoxLayout *m_buttonLayout;
    QPushButton *m_docToPdfButton;
    QPushButton *m_pdfToDocxButton;
    QProgressBar *m_progressBar;
    QLabel *m_statusLabel;
    QTextEdit *m_logOutput;
    
    // Converter
    DocPdf *m_converter;
    QThread *m_converterThread;
    
    QString m_currentDir;
};

#endif // MAINWINDOW_H