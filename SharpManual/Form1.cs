using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SharpManual
{
    public partial class Form1 : Form
    {
        private int RowId;

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


    }
}
