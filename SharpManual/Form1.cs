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

        private void OleDbDataAdapter2_RowUpdated(object sender, System.Data.OleDb.OleDbRowUpdatedEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet3.spaces". При необходимости она может быть перемещена или удалена.
            this.spacesTableAdapter.Fill(this.databaseDataSet3.spaces);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet2.classes". При необходимости она может быть перемещена или удалена.
            this.classesTableAdapter.Fill(this.databaseDataSet2.classes);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet1.methods". При необходимости она может быть перемещена или удалена.
            this.methodsTableAdapter.Fill(this.databaseDataSet1.methods);

        }
    }
}
