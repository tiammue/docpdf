#include "mainwindow.h"
#include <QApplication>
#include <QDir>
#include <QMessageBox>
#include <QFont>
#include <QStyle>
#include <QTime>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , m_centralWidget(nullptr)
    , m_converter(nullptr)
    , m_converterThread(nullptr)
{
    m_currentDir = getCurrentDirectory();
    setupUI();
    
    // Setup converter thread
    m_converterThread = new QThread(this);
    m_converter = new DocPdf();
    m_converter->moveToThread(m_converterThread);
    
    // Connect signals
    connect(m_converter, &DocPdf::progress, 
            this, &MainWindow::onConversionProgress);
    connect(m_converter, &DocPdf::finished, 
            this, &MainWindow::onConversionFinished);
    connect(m_converter, &DocPdf::error, 
            this, &MainWindow::onConversionError);
    
    m_converterThread->start();
}

MainWindow::~MainWindow()
{
    if (m_converterThread) {
        m_converterThread->quit();
        m_converterThread->wait();
    }
    delete m_converter;
}

void MainWindow::setupUI()
{
    setWindowTitle("docpdf");
    setFixedSize(500, 400);
    
    m_centralWidget = new QWidget(this);
    setCentralWidget(m_centralWidget);
    
    m_mainLayout = new QVBoxLayout(m_centralWidget);
    m_mainLayout->setSpacing(15);
    m_mainLayout->setContentsMargins(20, 20, 20, 20);
    
    // Title
    m_titleLabel = new QLabel("docpdf", this);
    QFont titleFont("Arial", 16, QFont::Bold);
    m_titleLabel->setFont(titleFont);
    m_titleLabel->setAlignment(Qt::AlignCenter);
    m_mainLayout->addWidget(m_titleLabel);
    
    // Directory info
    m_dirLabel = new QLabel(QString("Working Directory: %1").arg(QDir(m_currentDir).dirName()), this);
    QFont dirFont("Arial", 10);
    m_dirLabel->setFont(dirFont);
    m_dirLabel->setAlignment(Qt::AlignCenter);
    m_mainLayout->addWidget(m_dirLabel);
    
    // Buttons
    m_buttonLayout = new QHBoxLayout();
    
    m_docToPdfButton = new QPushButton("Convert DOC/DOCX → PDF", this);
    m_docToPdfButton->setMinimumHeight(60);
    m_docToPdfButton->setStyleSheet(
        "QPushButton {"
        "    background-color: #4CAF50;"
        "    color: white;"
        "    font-weight: bold;"
        "    font-size: 10pt;"
        "    border: none;"
        "    border-radius: 5px;"
        "}"
        "QPushButton:hover {"
        "    background-color: #45a049;"
        "}"
        "QPushButton:disabled {"
        "    background-color: #cccccc;"
        "}"
    );
    
    m_pdfToDocxButton = new QPushButton("Convert PDF → DOCX", this);
    m_pdfToDocxButton->setMinimumHeight(60);
    m_pdfToDocxButton->setStyleSheet(
        "QPushButton {"
        "    background-color: #2196F3;"
        "    color: white;"
        "    font-weight: bold;"
        "    font-size: 10pt;"
        "    border: none;"
        "    border-radius: 5px;"
        "}"
        "QPushButton:hover {"
        "    background-color: #1976D2;"
        "}"
        "QPushButton:disabled {"
        "    background-color: #cccccc;"
        "}"
    );
    
    connect(m_docToPdfButton, &QPushButton::clicked, this, &MainWindow::convertDocToPdf);
    connect(m_pdfToDocxButton, &QPushButton::clicked, this, &MainWindow::convertPdfToDocx);
    
    m_buttonLayout->addWidget(m_docToPdfButton);
    m_buttonLayout->addWidget(m_pdfToDocxButton);
    m_mainLayout->addLayout(m_buttonLayout);
    
    // Progress bar
    m_progressBar = new QProgressBar(this);
    m_progressBar->setVisible(false);
    m_mainLayout->addWidget(m_progressBar);
    
    // Status label
    m_statusLabel = new QLabel("Ready", this);
    QFont statusFont("Arial", 9);
    m_statusLabel->setFont(statusFont);
    m_statusLabel->setAlignment(Qt::AlignCenter);
    m_statusLabel->setStyleSheet("color: green;");
    m_mainLayout->addWidget(m_statusLabel);
    
    // Log output
    m_logOutput = new QTextEdit(this);
    m_logOutput->setMaximumHeight(100);
    m_logOutput->setReadOnly(true);
    m_logOutput->setFont(QFont("Consolas", 8));
    m_mainLayout->addWidget(m_logOutput);
}

QString MainWindow::getCurrentDirectory()
{
    return QApplication::applicationDirPath();
}

void MainWindow::updateStatus(const QString &message, const QString &color)
{
    m_statusLabel->setText(message);
    m_statusLabel->setStyleSheet(QString("color: %1;").arg(color));
    m_logOutput->append(QString("[%1] %2").arg(QTime::currentTime().toString()).arg(message));
}

void MainWindow::convertDocToPdf()
{
    m_docToPdfButton->setEnabled(false);
    m_pdfToDocxButton->setEnabled(false);
    m_progressBar->setVisible(true);
    m_progressBar->setValue(0);
    
    updateStatus("Converting DOC/DOCX to PDF...", "blue");
    
    QMetaObject::invokeMethod(m_converter, "convertDocToPdf", 
                             Qt::QueuedConnection,
                             Q_ARG(QString, m_currentDir));
}

void MainWindow::convertPdfToDocx()
{
    m_docToPdfButton->setEnabled(false);
    m_pdfToDocxButton->setEnabled(false);
    m_progressBar->setVisible(true);
    m_progressBar->setValue(0);
    
    updateStatus("Converting PDF to DOCX...", "blue");
    
    QMetaObject::invokeMethod(m_converter, "convertPdfToDocx", 
                             Qt::QueuedConnection,
                             Q_ARG(QString, m_currentDir));
}

void MainWindow::onConversionProgress(int current, int total, const QString &filename)
{
    m_progressBar->setMaximum(total);
    m_progressBar->setValue(current);
    updateStatus(QString("Converting %1/%2: %3").arg(current).arg(total).arg(filename), "blue");
}

void MainWindow::onConversionFinished(int converted, int total, const QString &type)
{
    m_progressBar->setVisible(false);
    m_docToPdfButton->setEnabled(true);
    m_pdfToDocxButton->setEnabled(true);
    
    QString message = QString("Conversion complete! %1/%2 %3 files converted").arg(converted).arg(total).arg(type);
    updateStatus(message, "green");
    
    QMessageBox::information(this, "Success", message);
}

void MainWindow::onConversionError(const QString &error)
{
    m_progressBar->setVisible(false);
    m_docToPdfButton->setEnabled(true);
    m_pdfToDocxButton->setEnabled(true);
    
    updateStatus("Conversion failed", "red");
    QMessageBox::critical(this, "Error", error);
}