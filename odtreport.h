#ifndef ODTREPORT_H
#define ODTREPORT_H

#include <QObject>
#include <QFile>
#include <QtXml>

class ODTReport : public QObject
{
    Q_OBJECT
public:
    explicit ODTReport (QObject *parent = 0);
    ODTReport (const QString &fileOpen, QObject *parent = 0);
    bool isReady() const;
    void setTemplate(QString fileTemplate);

public slots:
    /*!
     * \brief save сохранить итоговый документ
     * \param FolderToSave каталог для сохранения
     * \param fileName имя файла (без расширения), если имя файла не задано,
     * то присваивается текущая дата и время в формате "dd-MM-yyyy_hh-mm-ss"
     */
    void save(const QString FolderToSave, const QString &fileName = QString());
    /*!
     * \brief replaceText заменить текст. Заменяет столько раз сколько будет вызвана функция.
     * Если текст не найден ни чего не происходит.
     * \param oldValue прежнее значение.
     * \param newValue новое значение.
     */
    void replaceText(const QString &oldValue, const QString &newValue);
    /*!
     * \brief replaceCloneText многократная вставка текста на позицию искогомо текста.
     * Если текст не найден ни чего не происходит.
     * \param oldValue прежнее значение.
     * \param newValue новые значения.
     */
    void replaceCloneText(const QString &oldValue, const QStringList &newValue);
    /*!
     * \brief removeText удалить текст 'value'.
     * \param value текст для удаления. Если строка не является искомым текстом, но содержит его,
     * 'value' заменится на ''.
     */
    void removeText(const QString &value);
    /*!
     * \brief fillTemplateTable заполнить шаблонную таблицу. За шаблонную строку таблицы
     * принимается последняя строка.
     * \param TableName имя таблицы, которая будет заполняться.
     * \param oldText текст в шаблонной строке для замены QList <QString> oldText.
     * \param newText таблица в виде списка строк QList <QString> oldText. В каждой строке
     *  разделителем является "|".
     */
    void fillTemplateTable(const QString TableName,
                           QList <QString> oldText,
                           QList <QString> newText);
    /*!
     * \brief fillTable Заменить значения в таблице.
     * \param TableName имя таблицы, которая будет заполняться.
     * \param listVariable список искомых переменных.
     * \param listValue список значений.
     */
    void fillTable(const QString TableName,
                   const QList <QString> &listVariable,
                   const QList <QString> &listValue);
    /*!
     * \brief createTable создать таблицу.
     * \param table таблица в виде списка строк, значение ячеек в строке является QStringList.
     * \param textReaplaceOnTable текст в документе, куда будет вставлена новая таблица.
     * \param headerRow заголовок таблицы. Если параметр не задан, то заголовк будет заполнен
     * числами от 1 до количества значений в первой строке @see table
     */
    void createTable(QList <QStringList> table = QList <QStringList>(),
                     QString textReaplaceOnTable = QString(),
                     QStringList headerRow = QStringList());
    /*!
     * \brief replaceImage заменить изображение в шаблоне '/Pictures/' на @see fileName.
     * Замена происходит в алфавитном порядке имен в каталоге шаблона '/Pictures/'.
     * \param fileName полное имя файла изображения.
     */
    void replaceImage(const QString fileName);
    //**
    void replaceTextInTable(const QString tableName, const QString &Value, const QString &newValue);
    //**
    /*!
     * \brief addNewParagraph добавить текст в конец документа.
     * \param text значение текста.
     */
    void addNewParagraph(const QString &text);
signals:
    void documentIsDone();
private slots:
    void unZip(const QString &fileOpen);
    void addNewText(QDomDocument domDoc, QDomNode &node, const QString &Value);
    bool writeFile();
    QDomNode NeedTextNode(const QString &Value);
    QDomNode NeedTextInTableNode(QDomNode &node, const QString &Value);
    void replaceTextInRowTableNode(QDomNode &node, const QString &Value, const QString &newValue);
    QDomNode SetTextNode(QDomNode &node, const QString &newValue);
    QDomElement FindNeedTable(const QString TableName);
    /*!
     * \brief addStyleText добавление стилей текста для создания таблицы.
     * Вызов функции обязателен, если в документе нет других таблиц.
     */
    void addStyleText();
    /*!
     * \brief addStyleTable добавление стиля таблицы для создания таблицы.
     * Вызов функции обязателен, если в документе нет других таблиц.
     */
    void addStyleTable();
    /*!
     * \brief addStyleTableColumn добавление стиля колонок для создания таблицы.
     * Вызов функции обязателен, если в документе нет других таблиц.
     */
    void addStyleTableColumn();
    /*!
     * \brief addStyleTableCell добавление стилей ячеек для создания таблицы.
     * Вызов функции обязателен, если в документе нет других таблиц.
     */
    void addStyleTableCell();

private:
    bool m_isReady;///< Статус документа. При распаковке равен false.
    QString m_template; ///< Имя документа-шаблона.
    QByteArray m_textFile;
    QFile m_xmlFile;
    QDomDocument domDoc;
    QDomElement root;
    QDomNode MainNode;//основной текст
    QDomNode m_styleNode;

    QDomElement Table;///< Таблица.
    QDomNode TableRow;///< Строка таблицы шаблон.
    QDomNode CloneTableRow;///< Копия строки таблицы шаблона.

    QStringList m_imagesList;///< Список имен файлов изображений в шаблоне.
    int m_number_image;///< Номер изображения в списке, для замены.

};

#endif // ODTREPORT_H
