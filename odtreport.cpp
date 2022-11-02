#include "odtreport.h"

#include <QDir>
#include <QProcess>
#include <QMessageBox>


ODTReport::ODTReport(QObject *parent) :
    QObject(parent), m_isReady(false)
{
}

ODTReport::ODTReport(const QString &fileOpen, QObject *parent): QObject(parent), m_isReady(false)
{
    unZip(fileOpen);
}

void ODTReport::save(const QString FolderToSave, const QString &fileName)
{
    QDir dir;
    QProcess proc;
    QString tmpPath = "/tmp";
    QString reportPath_= FolderToSave;
    writeFile();
    proc.setWorkingDirectory( tmpPath );

    proc.execute( "bash -c \"cd " + tmpPath + "/shab ; zip " + tmpPath + "/shab.zip -r *\" " );
    proc.waitForFinished();
    QString outFileName;
    outFileName = QString( "%1/%2" ).arg( reportPath_ ).arg(
                QDateTime::currentDateTime().toString("dd-MM-yyyy_hh-mm-ss")+".odt");
    if (!fileName.isEmpty())
        outFileName = QString( "%1/%2" ).arg( reportPath_ ).arg(fileName+".odt");

    dir.remove( outFileName );
    if( dir.rename( tmpPath + "/shab.zip", outFileName ) )
    {
        QMessageBox msg;
        msg.setWindowTitle( " " );
        msg.setIcon( QMessageBox::Information );
        msg.setText( "<b>Документ готов.</b>" );
        msg.exec();
        m_isReady = true;
    }
    else
    {
        QMessageBox msg_critical;
        msg_critical.setWindowTitle( " " );
        msg_critical.setIcon(QMessageBox::Critical);
        msg_critical.setText("При формировании возникли ошибки.\nНевозможно записать в каталог назначения.\n"
                             "Проверьте правильность файла конфигурации и права на каталог назначения.");
    }
    proc.execute( "rm -rf " + tmpPath + "/shab" );

}

void ODTReport::replaceText(const QString &oldValue, const QString &newValue)
{
    QDomNode _node = NeedTextNode(oldValue);
    if (_node.toElement().text() != oldValue)
    {
        QString _str = _node.toElement().text();
        _str.replace(oldValue, newValue);
        MainNode.replaceChild(SetTextNode(_node, _str), _node);
    }
    else
    {
        MainNode.replaceChild(SetTextNode(_node, newValue), _node);
    }

}

void ODTReport::replaceCloneText(const QString &oldValue, const QStringList &newValue)
{
    QDomNode _node = NeedTextNode(oldValue);
    for (int i = 0; i < newValue.count(); i++)
    {
        QDomNode cloneNode = _node.cloneNode();
        if (cloneNode.toElement().text() != oldValue)
        {
            QString _str = cloneNode.toElement().text();
            _str.replace(oldValue, newValue.at(i));
            MainNode.insertBefore(cloneNode, _node);
            MainNode.replaceChild(SetTextNode(cloneNode, _str), cloneNode);
        }
        else
        {
            MainNode.insertBefore(cloneNode, _node);
            MainNode.replaceChild(SetTextNode(cloneNode, newValue.at(i)), cloneNode);
        }
    }
    MainNode.removeChild(_node);
}

void ODTReport::removeText(const QString &value)
{
    QDomNode _node = NeedTextNode(value);
    if (_node.toElement().text() != value)
    {
        QString _str = _node.toElement().text();
        _str.replace(value, "");
        MainNode.replaceChild(SetTextNode(_node, _str), _node);
    }
    else
    {
        MainNode.removeChild(_node);
    }
}

void ODTReport::replaceTextInRowTableNode(QDomNode &node, const QString &Value, const QString &newValue)
{
    QDomNode _node = node.firstChild();

    while (!_node.isNull())
    {
        QDomNode _res = NeedTextInTableNode(_node, Value);
        QString _str;
        if (_res.toElement().text() != Value && _res.toElement().text().indexOf(Value) != -1)
        {
            _str = _res.toElement().text();
            _str = _str.replace(Value, newValue);
            _res = SetTextNode(_res, _str);
        }
        else
        {
            _res = SetTextNode(_res, newValue);
        }
        _node = _node.nextSibling();
    }
}

void ODTReport::fillTemplateTable(const QString TableName, QList<QString> oldText, QList<QString> newText)
{
    Table = FindNeedTable(TableName);
    TableRow = Table.lastChildElement("table:table-row");
    for (int row = 0; row < newText.count(); row++)
    {
        CloneTableRow = TableRow.cloneNode();
        QStringList _list = newText[row].split('|');
        for (int col = 0; col < oldText.count(); col++)
        {
            replaceTextInRowTableNode(CloneTableRow, oldText[col],_list[col]);
        }
        Table.appendChild(CloneTableRow);
    }
    Table.removeChild(TableRow);
}

void ODTReport::fillTable(const QString TableName, const QList<QString> &listVariable, const QList<QString> &listValue)
{
    Table = FindNeedTable(TableName);
    for (int cell = 0; cell < listVariable.count(); cell++)
    {
        QDomNode Row = Table.firstChildElement("table:table-row");
        while ( !Row.isNull() )
        {
            replaceTextInRowTableNode(Row, listVariable.at(cell), listValue.at(cell));
            Row = Row.nextSiblingElement("table:table-row");
        }

    }

}


void ODTReport::createTable(QList<QStringList> table, QString textReaplaceOnTable,  QStringList headerRow)
{
    addStyleText();
    addStyleTable();
    addStyleTableColumn();
    addStyleTableCell();

    QList<QStringList> m_tab = table;

    QDomElement _table = domDoc.createElement("table:table");
    _table.setAttribute("table:name", "m_name");
    _table.setAttribute("table:style-name", "GeneratedTable");

    QDomElement _column = domDoc.createElement("table:table-column");
    _column.setAttribute("table:style-name", "GeneratedTable.A");

    _table.appendChild(_column);

    QStringList m_headerRow;
    if (headerRow.isEmpty())
    {
        _column.setAttribute("table:number-columns-repeated", QString::number(table[0].size()));
        for (int i = 0; i < table[0].size(); i++)
            m_headerRow.append(QString::number(i+1));
    }
    else
    {
        m_headerRow = headerRow;
        _column.setAttribute("table:number-columns-repeated", QString::number(headerRow.size()));
    }
    QDomElement _headersRow = domDoc.createElement("table:table-header-rows");
    QDomElement _headerRow = domDoc.createElement("table:table-row");

    for (int i = 0; i < m_headerRow.size(); i++)//заполнение заголовка
    {
        QDomElement cell = domDoc.createElement("table:table-cell");
        cell.setAttribute("table:style-name", "GeneratedTable.A1");
        cell.setAttribute("office:value-type", "string");

        QDomElement text = domDoc.createElement("text:p");
        text.setAttribute("text:style-name","THRow");

        QDomText t = domDoc.createTextNode(m_headerRow.at(i));

        text.appendChild(t);
        cell.appendChild(text);
        _headerRow.appendChild(cell);
    }

    _headersRow.appendChild(_headerRow);
    _table.appendChild(_headersRow);

    for (int i = 0; i < m_tab.size(); i++)
    {
        QStringList r = m_tab.at(i);
        QDomElement row = domDoc.createElement("table:table-row");
        if (r.at(0) == "&")
        {
            r.pop_front();
        }
        for (int j = 0; j < r.size(); j++)
        {
            QDomElement cell = domDoc.createElement("table:table-cell");

            cell.setAttribute("table:style-name", "GeneratedTable.A1");
            cell.setAttribute("office:value-type", "string");

            QDomElement text = domDoc.createElement("text:p");
            text.setAttribute("text:style-name", "GeneratedTable");

            QDomText t = domDoc.createTextNode(r.at(j));
            //если колонок в заголовоке больше
            if (j == (r.size()-1) && r.size() < m_headerRow.size())
            {
                text.appendChild(t);
                cell.appendChild(text);
                row.appendChild(cell);
                for (int diff = 0; diff < (m_headerRow.size() - r.size()); diff ++)
                {
                    QDomElement diff_cell = domDoc.toDocument().createElement("table:table-cell");
                    diff_cell.setAttribute("table:style-name", "GeneratedTable.A1");
                    diff_cell.setAttribute("office:value-type", "string");
                    row.appendChild(diff_cell);
                }

            }
            else
            {
                text.appendChild(t);
                cell.appendChild(text);
                row.appendChild(cell);
            }

        }
        _table.appendChild(row);
    }

    if (!NeedTextNode(textReaplaceOnTable).isNull())
        MainNode.replaceChild(_table, NeedTextNode(textReaplaceOnTable));
    else // если искомого слова не найдено
    {
        MainNode.replaceChild(_table, MainNode.lastChildElement("text:p"));
    }

}

void ODTReport::replaceImage(const QString fileName)
{
    if (m_number_image < m_imagesList.count())
    {
        QString oldImage = QString("/tmp/shab/Pictures/%1").arg(m_imagesList[m_number_image]);
        QFile::remove(oldImage);
        QFile::copy(fileName, oldImage);
    }
    m_number_image++;
}

void ODTReport::replaceTextInTable(const QString tableName, const QString &Value, const QString &newValue)
{
    Table = FindNeedTable(tableName);
    QDomNode tmpTableRow = Table.firstChildElement("table:table-row");
    bool res = false;
    while ( !tmpTableRow.isNull() )
    {
        QDomNode _node = tmpTableRow.firstChildElement("table:table-cell");

        while (!_node.isNull())
        {
            QDomNode _res = NeedTextInTableNode(_node, Value);
            if (!_res.isNull())
            {
                _res = SetTextNode(_res, newValue);
                res = true;
                break;
            }

            _node = _node.nextSiblingElement("table:table-cell");
        }
        if (res)
            break;
        tmpTableRow = tmpTableRow.nextSiblingElement("table:table-row");
    }

}

void ODTReport::addStyleText()
{

    QDomElement _style = domDoc.createElement("style:style");

    _style.setAttribute("style:family", "paragraph");
    _style.setAttribute("style:name", "THRow");
    _style.setAttribute("style:parent-style-name", "Table_20_Heading");

    QDomElement _style_paragraph_properties = domDoc.createElement("style:paragraph-properties");
    _style_paragraph_properties.setTagName("style:paragraph-properties");
    _style_paragraph_properties.setAttribute("fo:text-align", "center");
    _style_paragraph_properties.setAttribute("style:justify-single-word", "false");

    QDomElement _style_text_properties = domDoc.createElement("style:text-properties");
    _style_text_properties.setTagName("style:text-properties");
    _style_text_properties.setAttribute("style:font-weight-asian", "bold");
    _style_text_properties.setAttribute("style:font-weight-complex", "bold");

    _style.appendChild(_style_paragraph_properties);
    _style.appendChild(_style_text_properties);

    m_styleNode.appendChild(_style);
}

void ODTReport::addStyleTable()
{
    //***********************Таблица
    QDomElement _table_style = domDoc.createElement("style:style");
    _table_style.setTagName("style:style");

    _table_style.setAttribute("style:name", "GeneratedTable");
    _table_style.setAttribute("style:family", "table");

    QDomElement _table_style_properties = domDoc.createElement("style:table-properties");

    _table_style_properties.setTagName("style:table-properties");
    _table_style_properties.setAttribute("style:width", "17cm");
    _table_style_properties.setAttribute("table:align", "margins");

    _table_style.appendChild(_table_style_properties);

    m_styleNode.appendChild(_table_style);
}

void ODTReport::addStyleTableColumn()
{
    //***********************Колонка
    QDomElement _table_column_style = domDoc.createElement("style:style");
    _table_column_style.setTagName("style:style");
    _table_column_style.setAttribute("style:name", "GeneratedTable.A");
    _table_column_style.setAttribute("style:family", "table-column");

    QDomElement _table_column_style_properties = domDoc.createElement("style:table-column-properties");
    _table_column_style_properties.setTagName("style:table-column-properties");
    _table_column_style_properties.setAttribute("style:column-width", "16383*");
    _table_column_style_properties.setAttribute("style:rel-column-width", "16383*");

    _table_column_style.appendChild(_table_column_style_properties);

    m_styleNode.appendChild(_table_column_style);

}

void ODTReport::addStyleTableCell()
{
    //***********************Ячейка
    QDomElement _table_cell_style = domDoc.createElement("style:style");
    _table_cell_style.setTagName("style:style");

    _table_cell_style.setAttribute("style:name", "GeneratedTable.A1");
    _table_cell_style.setAttribute("style:family", "table-cell");

    QDomElement _table_cell_style_properties = domDoc.createElement("style:table-cell-properties");
    _table_cell_style_properties.setTagName("style:table-cell-properties");
    _table_cell_style_properties.setAttribute("fo:background-color", "transparent");
    _table_cell_style_properties.setAttribute("fo:padding", "0.097cm");
    _table_cell_style_properties.setAttribute("fo:border-left", "0.5pt solid #000000");
    _table_cell_style_properties.setAttribute("fo:border-right", "0.5pt solid #000000");
    _table_cell_style_properties.setAttribute("fo:border-top", "0.5pt solid #000000");
    _table_cell_style_properties.setAttribute("fo:border-bottom", "0.5pt solid #000000");
    _table_cell_style_properties.setAttribute("style:writing-mode", "page");

    _table_cell_style.appendChild(_table_cell_style_properties);

    m_styleNode.appendChild(_table_cell_style);

}


void ODTReport::addNewText(QDomDocument domDoc, QDomNode &node, const QString &Value)
{
    //--вставка своего текста
    QDomElement newNodeTag = domDoc.createElement(QString("text:p"));
    QDomText newNodeText = domDoc.createTextNode(QString(Value));
    newNodeTag.appendChild(newNodeText);
    node.appendChild(newNodeTag);
}

void ODTReport::addNewParagraph(const QString &text)
{
    //--вставка своего текста
    QDomElement newNodeTag = MainNode.toDocument().createElement(QString("text:p"));
    QDomText newNodeText = MainNode.toDocument().createTextNode(QString(text));
    newNodeTag.appendChild(newNodeText);
    MainNode.appendChild(newNodeTag);
}

void ODTReport::unZip(const QString &fileOpen)
{
    QDir dir;
    QProcess proc;
    m_xmlFile.setFileName( fileOpen );
    QString tmpPath = "/tmp";
    if( !m_xmlFile.exists() )
    {
        QMessageBox msgCritical;
        msgCritical.setWindowTitle( "Ошибка генерации файла" );
        msgCritical.setIcon( QMessageBox::Critical );
        msgCritical.setText( QString("<b>Отсутствует или нет доступа к файлу с шаблоном.</b>\n%1" )
                            .arg(fileOpen));
        msgCritical.exec();
        return;
    }
    proc.setWorkingDirectory( tmpPath );
    proc.execute( "rm -rf " + tmpPath + "/shab" );
    dir.mkpath( tmpPath + "/shab" );
    proc.execute( "unzip " + QFileInfo( m_xmlFile ).absoluteFilePath() + " -d " + tmpPath + "/shab" );
    proc.waitForFinished();
    m_xmlFile.setFileName(tmpPath + "/shab/content.xml");
    m_xmlFile.open( QIODevice::ReadOnly );
    m_textFile = m_xmlFile.readAll();
    m_xmlFile.close();

    domDoc.setContent(m_textFile);
    root = domDoc.documentElement();
    MainNode = root.firstChildElement("office:body").firstChild();//"office:text"
    m_styleNode = root.firstChildElement("office:automatic-styles");

    m_imagesList.clear();
    m_number_image = 0;
    if (QDir(tmpPath + "/shab/Pictures").exists())
    {
        m_imagesList = QDir(tmpPath + "/shab/Pictures").entryList(QDir::Files);
    }
}

bool ODTReport::writeFile()
{

    if (m_xmlFile.open(QIODevice::WriteOnly | QIODevice::Truncate))
    {
        QTextStream stream;
        stream.setDevice(&m_xmlFile);
        domDoc.save(stream,QDomNode::EncodingFromDocument);
        m_xmlFile.close();
        return true;
    }
    else
    {
        m_xmlFile.close();
        return false;
    }
}

QDomNode ODTReport::NeedTextNode(const QString &Value)
{
    QDomNode _find = MainNode.firstChildElement("text:p");
    while (!_find.isNull())
    {
        if(_find.toElement().text()==Value)
        {
            return _find;
            break;
        }
        else
        {
            if (_find.toElement().text().indexOf(Value) != -1)
            {
                QDomNode tmpSpan = _find.firstChildElement("text:span");
                while( !tmpSpan.isNull() )
                {
                    if (tmpSpan.toElement().text().indexOf(Value) != -1)
                    {
                        return tmpSpan;
                        break;
                    }
                    tmpSpan = tmpSpan.nextSiblingElement("text:span");
                }
                return _find;
                break;
            }
        }

        _find = _find.nextSiblingElement("text:p");
    }
    QDomNode emptyNode;
    return emptyNode;
}

QDomNode ODTReport::NeedTextInTableNode(QDomNode &node, const QString &Value)
{
    QDomNode _find = node.firstChildElement("text:p");
    while (!_find.isNull())
    {
        if(_find.toElement().text()==Value)
        {
            return _find;
            break;
        }
        else
        {
            if (_find.toElement().text().indexOf(Value) != -1)
            {
                QDomNode tmpSpan = _find.firstChildElement("text:span");
                while( !tmpSpan.isNull() )
                {
                    if (tmpSpan.toElement().text().indexOf(Value) != -1)
                    {
                        return tmpSpan;
                        break;
                    }
                    tmpSpan = tmpSpan.nextSiblingElement("text:span");
                }
                return _find;
                break;
            }
        }
        _find = _find.nextSiblingElement("text:p");
    }
    QDomNode emptyNode;
    return emptyNode;
}

QDomNode ODTReport::SetTextNode(QDomNode &node, const QString &newValue)
{
    QDomNode _result = node;
    _result.childNodes().item(0).setNodeValue(newValue);
    return _result;
}

QDomElement ODTReport::FindNeedTable(const QString TableName)
{
    QDomElement _result = MainNode.firstChildElement("table:table");
    while (!_result.isNull())
    {
        if (_result.tagName()=="table:table" && _result.attribute("table:name")==TableName)
            return _result;
        _result = _result.nextSiblingElement("table:table");
    }
    QDomElement emptyElement;
    return emptyElement;
}

bool ODTReport::isReady() const
{
    return m_isReady;
}

void ODTReport::setTemplate(QString fileTemplate)
{
    m_template = fileTemplate;
    unZip(m_template);
}
