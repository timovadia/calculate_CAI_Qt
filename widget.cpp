#include "widget.h"
#include "ui_widget.h"

#include <QApplication>
#include <QAxObject>
#include <QAxWidget>
#include <QtWidgets>
#include <QString>
#include <QStyleFactory>

Widget::Widget(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::Widget)
{
    ui->setupUi(this);

//    ui->buttonBox->button(QDialogButtonBox::Ok)->setText("Выполнить");
//    ui->buttonBox->button(QDialogButtonBox::Cancel)->setText("Отмена");

    ui->tabWidget->setCurrentWidget(ui->tab_DataXLS);//устанавливаем вкладку tab_DataXLS при запуске приложения

    QString sHelloEng = "Hello!";
    ui->fileName409->setReadOnly(true);
    ui->fileNameMarketRisk->setReadOnly(true);
    ui->fileNameODR->setReadOnly(true);
    ui->fileNameOSV->setReadOnly(true);
    ui->fileNameResultReport->setReadOnly(true);
    ui->listWidgetMessages->addItem(sHelloEng);

    ui->listWidgetMessages->setStyleSheet("QListWidget { background-color: gray }");
}

Widget::~Widget()
{
    delete ui;
}

void ExcelConnector(QString& file1, QString& file2, Ui::Widget& ui){

    ui.listWidgetMessages->addItem("Читаем файлы excel ...");

    try {

        QAxObject* excel = new QAxObject( "Excel.Application", 0 );
        QAxObject* workbooks = excel->querySubObject( "Workbooks" );
        QAxObject* workbookData = workbooks->querySubObject( "Open(const QString&)", file1 );
        QAxObject* workbookRezult = workbooks->querySubObject( "Open(const QString&)", file2 );
        QAxObject* sheet = workbookData->querySubObject( "Worksheets(int)", 1 );//выбираем первый лист книги
        QAxObject* sheetRez = workbookRezult->querySubObject( "Worksheets(int)", 1 );//выбираем первый лист книги

        //определяем число строк и столбцов
        QAxObject* usedRange = sheet->querySubObject("UsedRange");
        QAxObject* rows = usedRange->querySubObject("Rows");
        QAxObject* columns = usedRange->querySubObject("Columns");

        int countRows = rows->property("Count").toInt();
        int countCols = columns->property("Count").toInt();

        ui.listWidgetMessages->addItem("Чтение файлов прошло успешно.");

        QString sNumRows = "Число строк в xls файле равно: ";
        QString sNumCol = "Число столбцов в xls файле равно: ";

        ui.listWidgetMessages->addItem( sNumRows + QString::number(countRows) );
        ui.listWidgetMessages->addItem( sNumCol + QString::number(countCols) );

        // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
        int row = 5;
        int col = 1;
        QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row, col);
        // получение содержимого
        QVariant result = cell->property("Value");
        // освобождение памяти
        delete cell;

        // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
        QAxObject* cellRez = sheetRez->querySubObject("Cells(QVariant,QVariant)", 1, 1);
        // вставка значения переменной data (любой тип, приводимый к QVariant) в полученную ячейку
        cellRez->setProperty("Value", QVariant(result));
        // освобождение памяти
        delete cellRez;

        // очистка памяти
        delete columns;
        delete rows;
        delete usedRange;
        delete sheet;
        delete sheetRez;

        workbookData->dynamicCall("Close()");//close file
        workbookRezult->dynamicCall("Save()");
        workbookRezult->dynamicCall("Close()");
        excel->dynamicCall("Quit()");//close Excel
        delete workbookData;
        delete workbookRezult;
        delete workbooks;
        delete excel;

        ui.listWidgetMessages->addItem("Обработка файлов excel завершена!");

    }  catch (...) {

        ui.listWidgetMessages->addItem("Excel. Что-то пошло не так!");

    }

    ui.listWidgetMessages->addItem("The End ...");

}

void SendAlarmMessage(QString& text, Ui::Widget& ui) {
    QListWidgetItem* pItem =new QListWidgetItem(text);
    pItem->setForeground(Qt::red); // sets red text
    pItem->setBackground(Qt::green); // sets green background
    ui.listWidgetMessages->addItem(pItem);
//        ui->listWidget->show();

}

void SetDarkPalette() {
    qApp->setStyle(QStyleFactory::create("fusion"));

    QPalette darkPalette;

    // Настраиваем палитру для цветовых ролей элементов интерфейса
    darkPalette.setColor(QPalette::Window, QColor(53, 53, 53));
    darkPalette.setColor(QPalette::WindowText, Qt::white);
    darkPalette.setColor(QPalette::Base, QColor(25, 25, 25));
    darkPalette.setColor(QPalette::AlternateBase, QColor(53, 53, 53));
    darkPalette.setColor(QPalette::ToolTipBase, Qt::white);
    darkPalette.setColor(QPalette::ToolTipText, Qt::white);
    darkPalette.setColor(QPalette::Text, Qt::white);
    darkPalette.setColor(QPalette::Button, QColor(53, 53, 53));
    darkPalette.setColor(QPalette::ButtonText, Qt::white);
    darkPalette.setColor(QPalette::BrightText, Qt::red);
    darkPalette.setColor(QPalette::Link, QColor(42, 130, 218));
    darkPalette.setColor(QPalette::Highlight, QColor(42, 130, 218));
    darkPalette.setColor(QPalette::HighlightedText, Qt::black);

    // Устанавливаем данную палитру
    qApp->setPalette(darkPalette);
}

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    // Установим стиль оформления dark
    // Стандартная палитра является светлой
    SetDarkPalette();

    Widget mainWindow;
    mainWindow.setWindowTitle("Расчет ПДК. v.1.0.");



    mainWindow.show();
    return a.exec();
}


void Widget::on_buttonBox_accepted()
{
    if ( ui->fileNameOSV->text().isEmpty() || ui->fileNameODR->text().isEmpty() ||
         ui->fileName409->text().isEmpty() || ui->fileNameMarketRisk->text().isEmpty() ||
         ui->fileNameResultReport->text().isEmpty() ) {

         QString msg = "Не выбраны файлы!";
         SendAlarmMessage(msg, *ui);

    } else if ( ui->stringCompanyName->text().isEmpty() || ui->stringCompanyINN->text().isEmpty() ||
                ui->stringEmployee->text().isEmpty() || ui->stringEmployeeTel->text().isEmpty() ||
                ui->valueShareCapital->text().isEmpty() || ui->valueCapital->text().isEmpty() ||
                ui->valueReceivables->text().isEmpty() || ui->valueLossPreviousYears->text().isEmpty() ||
                ui->valueProfitCurrentYear->text().isEmpty() || ui->valueProfitPreviousYears->text().isEmpty() ) {

        QString msg = "Не заполнены показатели ручного ввода!";
        SendAlarmMessage(msg, *ui);
    }
    else {
        ui->listWidgetMessages->addItem("Ok!");
        QString fileData = ui->fileNameOSV->text();
        QString fileRez = ui->fileNameResultReport->text();
        ExcelConnector(fileData, fileRez, *ui);
    }

}

void Widget::on_buttonBox_rejected()
{
    this->close();
}

void Widget::on_toolButton_chooseFileOSV_clicked()
{
    QString fileOSV = QFileDialog::getOpenFileName(0, "Open Dialog", "", "*.xls *.xlsx");
    ui->fileNameOSV->setText(fileOSV);
}

void Widget::on_toolButton_chooseFileODR_clicked()
{
    QString fileODR = QFileDialog::getOpenFileName(0, "Open Dialog", "", "*.xls *.xlsx");
    ui->fileNameODR->setText(fileODR);
}

void Widget::on_toolButton_chooseFile409_clicked()
{
    QString file409 = QFileDialog::getOpenFileName(0, "Open Dialog", "", "*.xls *.xlsx");
    ui->fileName409->setText(file409);
}

void Widget::on_toolButton_chooseFileMarketRisk_clicked()
{
    QString fileMarketRisk = QFileDialog::getOpenFileName(0, "Open Dialog", "", "*.xls *.xlsx");
    ui->fileNameMarketRisk->setText(fileMarketRisk);
}

void Widget::on_toolButton_chooseFileResReport_clicked()
{
    QString fileResultReport = QFileDialog::getOpenFileName(0, "Open Dialog", "", "*.xls *.xlsx");
    ui->fileNameResultReport->setText(fileResultReport);
}
