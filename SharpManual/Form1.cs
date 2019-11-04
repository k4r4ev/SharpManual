using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace SharpManual
{
    public partial class Form1 : Form
    {
        private int RowId;
        Microsoft.Office.Interop.Word.Application wordApp;
        Microsoft.Office.Interop.Word.Document wordDoc;

        //объект приложения
        Microsoft.Office.Interop.Excel.Application ExcelApp;
        //объект окна Excel 
        Microsoft.Office.Interop.Excel.Window ExcelWindow;
        //объект рабочей книги
        Microsoft.Office.Interop.Excel.Workbook WorkBook;
        //набор листов Excel
        Microsoft.Office.Interop.Excel.Sheets ExcelSheets;
        //объект рабочего листа
        Microsoft.Office.Interop.Excel.Worksheet WorkSheet;
        //диапазон ячеек
        Microsoft.Office.Interop.Excel.Range range;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet3.spaces". При необходимости она может быть перемещена или удалена.
            this.spacesTableAdapter.Fill(this.databaseDataSet3.spaces);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet2.classes". При необходимости она может быть перемещена или удалена.
            this.classesTableAdapter.Fill(this.databaseDataSet2.classes);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet1.methods". При необходимости она может быть перемещена или удалена.
            this.methodsTableAdapter.Fill(this.databaseDataSet1.methods);

            oleDbDataAdapter1.Fill(databaseDataSet.methods);
            oleDbDataAdapter1.Fill(databaseDataSet.classes);
            oleDbDataAdapter1.Fill(databaseDataSet.spaces);

            oleDbConnection1.Open(); //открыть соединение

            UpdateTables();        //обновить главное окно
        }

        private void UpdateTables()
        {
            //обновить содержимое таблиц базы данных
            this.spacesTableAdapter.Update(this.databaseDataSet3.spaces);
            this.classesTableAdapter.Update(this.databaseDataSet2.classes);
            this.methodsTableAdapter.Update(this.databaseDataSet1.methods);
            //очистить DataSet
            databaseDataSet1.Clear();
            databaseDataSet2.Clear();
            databaseDataSet3.Clear();
            //заполнить таблицы в объекте 
            this.spacesTableAdapter.Fill(this.databaseDataSet3.spaces);
            this.classesTableAdapter.Fill(this.databaseDataSet2.classes);
            this.methodsTableAdapter.Fill(this.databaseDataSet1.methods);

            //очистить содерижмое списка классов
            comboBox1.Items.Clear();
            //заполнить список классов значениями из таблицы
            foreach (DataRow row in databaseDataSet2.classes)
                comboBox1.Items.Add(row["Название"]);
            //в списке классов - ни один класс не выделен
            comboBox1.Text = "";

            //очистить содерижмое списка пространств
            comboBox2.Items.Clear();
            //заполнить список пространств значениями из таблицы
            foreach (DataRow row in databaseDataSet3.spaces)
                comboBox2.Items.Add(row["Пространство"]);
            //в списке пространств - ни одно пространство не выделено
            comboBox2.Text = "";
        }





        // METHODS
        private void AddMethodButton_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "" && comboBox1.Text != "")
            {
                //если текстовые поля не пусты,
                try
                {
                    //то создать новую запись в таблице Contacts,
                    DataRow row = databaseDataSet1.methods.NewRow();
                    //заполнить ее столбцы
                    row["Название"] = textBox1.Text;
                    row["Описание"] = textBox2.Text;
                    //получить из выпадающего списка классы,
                    string selectedClass = comboBox1.SelectedItem.ToString();
                    //составить условие для поиска этого класса
                    string str = "Название='" + selectedClass + "'";
                    //получить id этого класса в таблице classes
                    DataRow[] classes = databaseDataSet2.classes.Select(str);
                    row["Id_класса"] = classes[0]["Код"];
                    //добавить запись в таблицу
                    databaseDataSet1.methods.Rows.Add(row);
                    //обновить форму
                    UpdateTables();
                }
                catch (Exception) { }
            }
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                //получить номер выделенной строки
                RowId = e.RowIndex;
                //отобразить фамилию и имя выбранного человека 
                //в текстовых полях
                textBox1.Text = databaseDataSet1.methods.Rows[RowId]["Название"].ToString();
                textBox2.Text = databaseDataSet1.methods.Rows[RowId]["Описание"].ToString();

                string str = "Код='" + databaseDataSet1.methods.Rows[RowId]["Id_класса"] + "'";
                //получить id этого класса в таблице classes
                DataRow[] classes = databaseDataSet2.classes.Select(str);
                comboBox1.Text = classes[0]["Название"].ToString();
            }
            catch (Exception) { }
        }

        private void dataGridView1_EditClick(object sender, EventArgs e)
        {
            //если выбрана строка в таблице
            if (dataGridView1.SelectedRows.Count != 0)
            {
                //получить содержимое выбранной строки
                DataRow row = databaseDataSet1.methods.Rows[RowId];
                //изменить фамилию и имя на введенные значения
                row["Название"] = textBox1.Text;
                row["Описание"] = textBox2.Text;
                string selectedClass = comboBox1.SelectedItem.ToString();
                //составить условие для поиска этого класса
                string str = "Название='" + selectedClass + "'";
                //получить id этого класса в таблице classes
                DataRow[] classes = databaseDataSet2.classes.Select(str);
                row["Id_класса"] = classes[0]["Код"];
                //сохранить изменения и обновить содержимое формы
                UpdateTables();
            }
            else
                MessageBox.Show("Выберите строку для редактирования", "Ошибка");
        }

        private void DelContactButton_Click(object sender, EventArgs e)
        {
            //если выбрана запись для удаления
            if (dataGridView1.SelectedRows.Count != 0)
            {
                if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        //удалить выбранную строку
                        databaseDataSet1.methods.Rows[RowId].Delete();
                    }
                    catch (Exception)
                    {
                    }
                    //обновить БД и ее содержимое на форме
                    UpdateTables();
                }
            }
            else
                MessageBox.Show("Выберите строку для удаления", "Ошибка");
        }






        // CLASSES
        private void AddClassesButton_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "" && textBox4.Text != "" && comboBox2.Text != "")
            {
                //если текстовые поля не пусты,
                try
                {
                    //то создать новую запись в таблице Contacts,
                    DataRow row = databaseDataSet2.classes.NewRow();
                    //заполнить ее столбцы
                    row["Название"] = textBox4.Text;
                    row["Описание"] = textBox3.Text;
                    //получить из выпадающего списка классы,
                    string selectedSpaces = comboBox2.SelectedItem.ToString();
                    //составить условие для поиска этого класса
                    string str = "Пространство='" + selectedSpaces + "'";
                    //получить id этого класса в таблице classes
                    DataRow[] classes = databaseDataSet3.spaces.Select(str);
                    row["Id_пространства"] = classes[0]["Код"];
                    //добавить запись в таблицу
                    databaseDataSet2.classes.Rows.Add(row);
                    //обновить форму
                    UpdateTables();
                }
                catch (Exception) { }
            }
        }

        private void dataGridView2_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                //получить номер выделенной строки
                RowId = e.RowIndex;
                //отобразить фамилию и имя выбранного человека 
                //в текстовых полях
                textBox4.Text = databaseDataSet2.classes.Rows[RowId]["Название"].ToString();
                textBox3.Text = databaseDataSet2.classes.Rows[RowId]["Описание"].ToString();

                string str = "Код='" + databaseDataSet2.classes.Rows[RowId]["Id_пространства"] + "'";
                //получить id этого класса в таблице classes
                DataRow[] spaces = databaseDataSet3.spaces.Select(str);
                comboBox2.Text = spaces[0]["Пространство"].ToString();
            }
            catch (Exception) { }
        }

        private void dataGridView2_EditClick(object sender, EventArgs e)
        {
            //если выбрана строка в таблице
            if (dataGridView2.SelectedRows.Count != 0)
            {
                //получить содержимое выбранной строки
                DataRow row = databaseDataSet2.classes.Rows[RowId];
                //изменить фамилию и имя на введенные значения
                row["Название"] = textBox4.Text;
                row["Описание"] = textBox3.Text;
                string selectedSpaces = comboBox2.SelectedItem.ToString();
                //составить условие для поиска этого класса
                string str = "Пространство='" + selectedSpaces + "'";
                //получить id этого класса в таблице classes
                DataRow[] spaces = databaseDataSet3.spaces.Select(str);
                row["Id_пространства"] = spaces[0]["Код"];
                //сохранить изменения и обновить содержимое формы
                UpdateTables();
            }
            else
                MessageBox.Show("Выберите строку для редактирования", "Ошибка");
        }

        private void DelClassesButton_Click(object sender, EventArgs e)
        {
            //если выбрана запись для удаления
            if (dataGridView2.SelectedRows.Count != 0)
            {
                if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        //удалить выбранную строку
                        databaseDataSet2.classes.Rows[RowId].Delete();
                    }
                    catch (Exception)
                    {
                    }
                    //обновить БД и ее содержимое на форме
                    UpdateTables();
                }
            }
            else
                MessageBox.Show("Выберите строку для удаления", "Ошибка");
        }






        // SPACES
        private void AddSpaceButton_Click(object sender, EventArgs e)
        {
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                //если текстовые поля не пусты,
                try
                {
                    //то создать новую запись в таблице Contacts,
                    DataRow row = databaseDataSet3.spaces.NewRow();
                    //заполнить ее столбцы
                    row["Пространство"] = textBox6.Text;
                    row["Описание"] = textBox5.Text;
                    databaseDataSet3.spaces.Rows.Add(row);
                    //обновить форму
                    UpdateTables();
                }
                catch (Exception) { }
            }
        }
        private void dataGridView3_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                //получить номер выделенной строки
                RowId = e.RowIndex;
                //отобразить фамилию и имя выбранного человека 
                //в текстовых полях
                textBox6.Text = databaseDataSet3.spaces.Rows[RowId]["Пространство"].ToString();
                textBox5.Text = databaseDataSet3.spaces.Rows[RowId]["Описание"].ToString();
            }
            catch (Exception) { }
        }

        private void dataGridView3_EditClick(object sender, EventArgs e)
        {
            //если выбрана строка в таблице
            if (dataGridView3.SelectedRows.Count != 0)
            {
                //получить содержимое выбранной строки
                DataRow row = databaseDataSet3.spaces.Rows[RowId];
                //изменить фамилию и имя на введенные значения
                row["Пространство"] = textBox6.Text;
                row["Описание"] = textBox5.Text;
                UpdateTables();
            }
            else
                MessageBox.Show("Выберите строку для редактирования", "Ошибка");
        }

        private void DelSpacesButton_Click(object sender, EventArgs e)
        {
            //если выбрана запись для удаления
            if (dataGridView3.SelectedRows.Count != 0)
            {
                if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        //удалить выбранную строку
                        databaseDataSet3.spaces.Rows[RowId].Delete();
                    }
                    catch (Exception) { }
                    //обновить БД и ее содержимое на форме
                    UpdateTables();
                }
            }
            else
                MessageBox.Show("Выберите строку для удаления", "Ошибка");
        }
        private void FilterCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            //если установлен флажок и выбрана фамилия
            if (checkBox1.Checked && textBox4.Text != "")
            {
                //получить пространство
                string selectedClass = textBox4.Text;
                //составить условие для поиска нужного человека 
                //в таблице Contacts
                string str = "Название='" + selectedClass + "'";
                //найти нужного человека в таблице Contacts
                DataRow[] classes = databaseDataSet2.classes.Select(str);
                //составить условие для фильтра
                str = "Id_класса=" + classes[0]["Код"];
                //применить фильтр
                methodsBindingSource.Filter = str;
            }
            else
                //отменить фильтрацию
                methodsBindingSource.Filter = "";
        }





        // WORD
        private void OpenDocument(string FileName)
        {
            //открываем Word
            wordApp = new Microsoft.Office.Interop.Word.Application();
            //создаем документ на основе шаблона
            Object template = System.Windows.Forms.Application.StartupPath + @"\docs\" + FileName;
            Object newTemplate = false;
            Object documentType = Microsoft.Office.Interop.Word.WdNewDocumentType.wdNewBlankDocument;
            Object visible = true;
            //добавляем документ в список документов приложения
            wordDoc = wordApp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
        }
        public void ReplaceText(string word, string repl)
        {
            // Смещаем выделение к началу документа
            Object unit = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            Object extend = Microsoft.Office.Interop.Word.WdMovementType.wdMove;
            wordApp.Selection.HomeKey(ref unit, ref extend);
            //создаем объект Find для поиска текста
            Microsoft.Office.Interop.Word.Find fnd = wordApp.Selection.Find;
            //очищаем его настройки
            fnd.ClearFormatting();
            //задаем текст для поиска
            fnd.Text = word;
            //очищаем настройки для замены
            fnd.Replacement.ClearFormatting();
            //задаем текст для замены
            fnd.Replacement.Text = repl;
            //запускаем процесс поиска и замены
            ExecuteReplace(fnd);
        }
        private Boolean ExecuteReplace(Microsoft.Office.Interop.Word.Find find)
        {
            return ExecuteReplace(find, Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }
        private Boolean ExecuteReplace(Microsoft.Office.Interop.Word.Find find, Object replaceOption)
        {
            Object findText = Type.Missing;
            Object matchCase = Type.Missing;
            Object matchWholeWord = Type.Missing;
            Object matchWildcards = Type.Missing;
            Object matchSoundsLike = Type.Missing;
            Object matchAllWordForms = Type.Missing;
            Object forward = Type.Missing;
            Object wrap = Type.Missing;
            Object format = Type.Missing;
            Object replaceWith = Type.Missing;
            Object replace = replaceOption;
            Object matchKashida = Type.Missing;
            Object matchDiacritics = Type.Missing;
            Object matchAlefHamza = Type.Missing;
            Object matchControl = Type.Missing;

            return find.Execute(ref findText, ref matchCase,
            ref matchWholeWord, ref matchWildcards, ref matchSoundsLike,
            ref matchAllWordForms, ref forward, ref wrap, ref format,
            ref replaceWith, ref replace, ref matchKashida,
            ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
        private void InvitationButton_Click(object sender, EventArgs e)
        {
            //если выбран человек для формирования приглашения
            if (dataGridView1.SelectedRows.Count != 0)
            {
                //создаем форму для ввода доп. информации
                Form2 form = new Form2();
                form.button1.DialogResult = DialogResult.OK;
                //если мы ввели данные и нажали ОК, то формируем документ
                if (form.ShowDialog() == DialogResult.OK)
                {
                    //создаем новый документ на основе шаблона
                    OpenDocument("Document1.docx");
                    //получаем из БД строку с выбранным человеком
                    DataRow row = databaseDataSet1.methods.Rows[RowId];
                    //получаем его имя и фамилию
                    //string FIO = row["Name"].ToString() + " " + row["Fam"].ToString();
                    //заменяем метки в шаблоне конкретными значениями
                    ReplaceText("<project>", form.textBox1.Text);
                    ReplaceText("<version>", form.textBox2.Text);
                    ReplaceText("<method>", textBox1.Text);
                    ReplaceText("<properties>", textBox2.Text);
                    //делаем приложение Word видимым
                    wordApp.Visible = true;
                }
            }
            else
                MessageBox.Show("Выберите метод", "Ошибка");
        }
        private void NumbersButton_Click(object sender, EventArgs e)
        {
            //создаем новый документ на основе шаблона
            OpenDocument("Document2.docx");
            //заменяем метку <Today> на текущую дату
            ReplaceText("<Today>", DateTime.Today.ToShortDateString());
            //задаем параметры для поиска метки <Table>
            Object start = 0;
            Object end = wordDoc.Characters.Count;
            //диапазон поиска - весь документ
            Microsoft.Office.Interop.Word.Range rng = wordDoc.Range(ref start, ref end);
            rng.TextRetrievalMode.IncludeHiddenText = false;
            rng.TextRetrievalMode.IncludeFieldCodes = false;
            string metka = "<Table>";
            //ищем в документе метку <Table>
            int beginphrase = rng.Text.IndexOf(metka);
            //получаем "координаты" начала и конца метки в документе
            start = beginphrase;
            end = beginphrase + metka.Length;
            //если метка <Table> найдена
            if (beginphrase != -1)
            {
                //то удаляем ее
                rng = wordDoc.Range(ref start, ref end);
                rng.Text = "";

                //и вставляем на ее место таблицу
                Object defaultTableBehavior = Type.Missing;
                Object autoFitBehavior = Type.Missing;
                //создаем объект таблицы (изначально - только шапка)
                Microsoft.Office.Interop.Word.Table tbl = rng.Tables.Add(rng, 1, 3, ref defaultTableBehavior, ref autoFitBehavior);

                //Форматируем таблицу и применяем стиль
                tbl.Range.Font.Size = 14;
                Object style = "Сетка таблицы";
                tbl.set_Style(ref style);

                //шапка таблицы
                tbl.Cell(1, 1).Range.Text = "Код";
                tbl.Cell(1, 2).Range.Text = "Метод";
                tbl.Cell(1, 3).Range.Text = "Описание";

                //i - общее количество строк в формируемой таблице
                int i = 0;
                //перебираем всех людей из таблицы БД Contacts
                foreach (DataRow row in databaseDataSet1.methods)
                {
                    i++;
                    //добавляем в таблицу документа новую строку
                    Object beforeRow = Type.Missing;
                    tbl.Rows.Add(ref beforeRow);
                    //и заполняем ее столбцы
                    tbl.Cell(i + 1, 1).Range.Text = i.ToString();
                    tbl.Cell(i + 1, 2).Range.Text = row["Название"].ToString();
                    tbl.Cell(i + 1, 3).Range.Text = row["Описание"].ToString();
                }
                //шапку таблицы выделяем курсивом
                tbl.Rows[1].Range.Font.Italic = 1;
                //и устанавливаем выравнивание по центру
                tbl.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            else
                ReplaceText("Table", "");
            //отображаем сформированный документ
            wordApp.Visible = true;
        }





        // EXCEL
        private void OpenExcelDocument(string FileName)
        {
            //создать новый объект приложения Excel
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            //задать файл шаблона
            Object template = System.Windows.Forms.Application.StartupPath + @"\docs\" + FileName;
            //применить шаблон
            ExcelApp.Workbooks.Add(template);
            //получить первую рабочую книгу файла
            WorkBook = ExcelApp.Workbooks[1];
            //получить список листов рабочей книги
            ExcelSheets = WorkBook.Worksheets;
            //выбрать первый лист
            WorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelSheets.get_Item(1);
        }
        private void PutCell(string cell, string val)
        {
            //получить диапазон, соответствующий выбранной ячейке
            range = WorkSheet.get_Range(cell, Type.Missing);
            //занести в ячейку значение
            range.Value2 = val;
        }
        private void PutCellBorder(string cell, string val)
        {
            //вызвать функцию занесения в ячейку значения
            PutCell(cell, val);
            //нарисовать границу вокруг ячейки
            range.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
        }
        private void ExcelButton_Click(object sender, EventArgs e)
        {
            //создать документ на основе шаблона
            OpenExcelDocument("Sheets1.xlsx");
            //занести текущую дату в ячейку D1
            PutCell("D1", DateTime.Now.ToShortDateString());
            //i - порядковый номер записи
            int i = 1;
            //просмотреть все строки таблицы Methods
            foreach (DataRow row in databaseDataSet1.methods)
            {
                PutCellBorder("B" + (i + 5).ToString(), row["Код"].ToString());
                PutCellBorder("C" + (i + 5).ToString(), row["Название"].ToString());
                PutCellBorder("D" + (i + 5).ToString(), row["Описание"].ToString());
                i++;
            }

            //сделать приложение Excel видимым
            ExcelApp.Visible = true;
        }
        private void ExcelButton2_Click(object sender, EventArgs e)
        {
            //создать документ на основе шаблона
            OpenExcelDocument("Sheets1.xlsx");
            //занести текущую дату в ячейку D1
            PutCell("D1", DateTime.Now.ToShortDateString());
            int i = 6;
            //просмотреть все строки таблицы Classes
            foreach (DataRow row in databaseDataSet.classes)
            {
                //MessageBox.Show("Выберите метод", "Ошибка");
                //занести в столбец А название класса
                PutCell("A" + i.ToString(), row["Название"].ToString());
                //выделить ячейки с А по D
                range = WorkSheet.get_Range("A" + i.ToString(), "C" + i.ToString());
                //и объединить их
                range.Merge(Type.Missing);
                //установить жирный шрифт
                range.Font.Bold = true;
                //и курсив
                range.Font.Italic = true;
                //нарисовать границу вокруг ячейки
                range.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                //установить выравнивание по центру
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //получить список телефонов по заданному классу
                DataRow[] methods = row.GetChildRows(databaseDataSet.Relations["КлассыМетоды"]);
                i++;
                //num - порядковый номер метода
                int num = 0;
                //просмотреть все методы класса
                foreach (DataRow method in methods)
                {
                    num++;
                    //занести в столбец А порядковый номер записи
                    PutCellBorder("A" + i.ToString(), num.ToString());
                    PutCellBorder("B" + i.ToString(), method["Название"].ToString());
                    PutCellBorder("C" + i.ToString(), method["Описание"].ToString());
                    i++;
                }
                //вывести количество записей в текущей группе
                PutCell("A" + i.ToString(), "Итого: " + num.ToString());
                //и отформатировать соответствующую ячейку
                range = WorkSheet.get_Range("A" + i.ToString(), "C" + i.ToString());
                range.Merge(Type.Missing);
                range.Font.Italic = true;
                range.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                i++;
            }
            //сделать приложение Excel видимым
            ExcelApp.Visible = true;
        }
    }
}
