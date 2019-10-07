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
            oleDbDataAdapter1.Update(databaseDataSet);
            oleDbDataAdapter2.Update(databaseDataSet);
            oleDbDataAdapter3.Update(databaseDataSet);
            //очистить DataSet
            databaseDataSet.Clear();
            //заполнить таблицы в объекте 
            oleDbDataAdapter1.Fill(databaseDataSet.methods);
            oleDbDataAdapter2.Fill(databaseDataSet.classes);
            oleDbDataAdapter3.Fill(databaseDataSet.spaces);

            //очистить содерижмое списка классов
            comboBox1.Items.Clear();
            //заполнить список классов значениями из таблицы
            foreach (DataRow row in databaseDataSet.classes)
                comboBox1.Items.Add(row["Название"]);
            //в списке классов - ни один класс не выделен
            comboBox1.Text = "";

            //очистить содерижмое списка пространств
            comboBox2.Items.Clear();
            //заполнить список пространств значениями из таблицы
            foreach (DataRow row in databaseDataSet.spaces)
                comboBox2.Items.Add(row["Пространство"]);
            //в списке пространств - ни одно пространство не выделено
            comboBox2.Text = "";
        }
    }
}
