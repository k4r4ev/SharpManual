namespace SharpManual
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.databaseDataSet = new SharpManual.DatabaseDataSet();
            this.oleDbConnection1 = new System.Data.OleDb.OleDbConnection();
            this.oleDbSelectCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDataAdapter1 = new System.Data.OleDb.OleDbDataAdapter();
            this.oleDbSelectCommand2 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDataAdapter2 = new System.Data.OleDb.OleDbDataAdapter();
            this.oleDbSelectCommand3 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDataAdapter3 = new System.Data.OleDb.OleDbDataAdapter();
            this.oleDbInsertCommand = new System.Data.OleDb.OleDbCommand();
            this.oleDbUpdateCommand = new System.Data.OleDb.OleDbCommand();
            this.oleDbDeleteCommand = new System.Data.OleDb.OleDbCommand();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.databaseDataSet1 = new SharpManual.DatabaseDataSet1();
            this.methodsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.methodsTableAdapter = new SharpManual.DatabaseDataSet1TableAdapters.methodsTableAdapter();
            this.databaseDataSet2 = new SharpManual.DatabaseDataSet2();
            this.classesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.classesTableAdapter = new SharpManual.DatabaseDataSet2TableAdapters.classesTableAdapter();
            this.databaseDataSet3 = new SharpManual.DatabaseDataSet3();
            this.spacesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.spacesTableAdapter = new SharpManual.DatabaseDataSet3TableAdapters.spacesTableAdapter();
            this.кодDataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.пространствоDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.описаниеDataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.кодDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.названиеDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.описаниеDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.idклассаDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.кодDataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.названиеDataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.описаниеDataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.idпространстваDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.methodsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.classesBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.spacesBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // databaseDataSet
            // 
            this.databaseDataSet.DataSetName = "DatabaseDataSet";
            this.databaseDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // oleDbConnection1
            // 
            this.oleDbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\karae\\Documents\\SharpManual" +
    "\\SharpManual\\Database.mdb";
            // 
            // oleDbSelectCommand1
            // 
            this.oleDbSelectCommand1.CommandText = "SELECT * FROM methods";
            this.oleDbSelectCommand1.Connection = this.oleDbConnection1;
            // 
            // oleDbDataAdapter1
            // 
            this.oleDbDataAdapter1.SelectCommand = this.oleDbSelectCommand1;
            this.oleDbDataAdapter1.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "methods", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("Код", "Код"),
                        new System.Data.Common.DataColumnMapping("Название", "Название"),
                        new System.Data.Common.DataColumnMapping("Описание", "Описание"),
                        new System.Data.Common.DataColumnMapping("Id_класса", "Id_класса")})});
            // 
            // oleDbSelectCommand2
            // 
            this.oleDbSelectCommand2.CommandText = "SELECT * FROM classes";
            this.oleDbSelectCommand2.Connection = this.oleDbConnection1;
            // 
            // oleDbDataAdapter2
            // 
            this.oleDbDataAdapter2.DeleteCommand = this.oleDbDeleteCommand;
            this.oleDbDataAdapter2.InsertCommand = this.oleDbInsertCommand;
            this.oleDbDataAdapter2.SelectCommand = this.oleDbSelectCommand2;
            this.oleDbDataAdapter2.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "classes", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("Код", "Код"),
                        new System.Data.Common.DataColumnMapping("Название", "Название"),
                        new System.Data.Common.DataColumnMapping("Описание", "Описание"),
                        new System.Data.Common.DataColumnMapping("Id_пространства", "Id_пространства")})});
            this.oleDbDataAdapter2.UpdateCommand = this.oleDbUpdateCommand;
            this.oleDbDataAdapter2.RowUpdated += new System.Data.OleDb.OleDbRowUpdatedEventHandler(this.OleDbDataAdapter2_RowUpdated);
            // 
            // oleDbSelectCommand3
            // 
            this.oleDbSelectCommand3.CommandText = "SELECT * FROM spaces";
            this.oleDbSelectCommand3.Connection = this.oleDbConnection1;
            // 
            // oleDbDataAdapter3
            // 
            this.oleDbDataAdapter3.SelectCommand = this.oleDbSelectCommand3;
            this.oleDbDataAdapter3.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "spaces", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("Код", "Код"),
                        new System.Data.Common.DataColumnMapping("Пространство", "Пространство"),
                        new System.Data.Common.DataColumnMapping("Описание", "Описание")})});
            // 
            // oleDbInsertCommand
            // 
            this.oleDbInsertCommand.CommandText = "INSERT INTO `classes` (`Название`, `Описание`, `Id_пространства`) VALUES (?, ?, ?" +
    ")";
            this.oleDbInsertCommand.Connection = this.oleDbConnection1;
            this.oleDbInsertCommand.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Название", System.Data.OleDb.OleDbType.VarWChar, 0, "Название"),
            new System.Data.OleDb.OleDbParameter("Описание", System.Data.OleDb.OleDbType.VarWChar, 0, "Описание"),
            new System.Data.OleDb.OleDbParameter("Id_пространства", System.Data.OleDb.OleDbType.Integer, 0, "Id_пространства")});
            // 
            // oleDbUpdateCommand
            // 
            this.oleDbUpdateCommand.CommandText = resources.GetString("oleDbUpdateCommand.CommandText");
            this.oleDbUpdateCommand.Connection = this.oleDbConnection1;
            this.oleDbUpdateCommand.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Название", System.Data.OleDb.OleDbType.VarWChar, 0, "Название"),
            new System.Data.OleDb.OleDbParameter("Описание", System.Data.OleDb.OleDbType.VarWChar, 0, "Описание"),
            new System.Data.OleDb.OleDbParameter("Id_пространства", System.Data.OleDb.OleDbType.Integer, 0, "Id_пространства"),
            new System.Data.OleDb.OleDbParameter("Original_Код", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Код", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("IsNull_Название", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "Название", System.Data.DataRowVersion.Original, true, null),
            new System.Data.OleDb.OleDbParameter("Original_Название", System.Data.OleDb.OleDbType.VarWChar, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Название", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("IsNull_Описание", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "Описание", System.Data.DataRowVersion.Original, true, null),
            new System.Data.OleDb.OleDbParameter("Original_Описание", System.Data.OleDb.OleDbType.VarWChar, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Описание", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("IsNull_Id_пространства", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "Id_пространства", System.Data.DataRowVersion.Original, true, null),
            new System.Data.OleDb.OleDbParameter("Original_Id_пространства", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Id_пространства", System.Data.DataRowVersion.Original, null)});
            // 
            // oleDbDeleteCommand
            // 
            this.oleDbDeleteCommand.CommandText = resources.GetString("oleDbDeleteCommand.CommandText");
            this.oleDbDeleteCommand.Connection = this.oleDbConnection1;
            this.oleDbDeleteCommand.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Original_Код", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Код", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("IsNull_Название", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "Название", System.Data.DataRowVersion.Original, true, null),
            new System.Data.OleDb.OleDbParameter("Original_Название", System.Data.OleDb.OleDbType.VarWChar, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Название", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("IsNull_Описание", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "Описание", System.Data.DataRowVersion.Original, true, null),
            new System.Data.OleDb.OleDbParameter("Original_Описание", System.Data.OleDb.OleDbType.VarWChar, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Описание", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("IsNull_Id_пространства", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "Id_пространства", System.Data.DataRowVersion.Original, true, null),
            new System.Data.OleDb.OleDbParameter("Original_Id_пространства", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Id_пространства", System.Data.DataRowVersion.Original, null)});
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.кодDataGridViewTextBoxColumn,
            this.названиеDataGridViewTextBoxColumn,
            this.описаниеDataGridViewTextBoxColumn,
            this.idклассаDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.methodsBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(757, 196);
            this.dataGridView1.TabIndex = 0;
            // 
            // dataGridView2
            // 
            this.dataGridView2.AutoGenerateColumns = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.кодDataGridViewTextBoxColumn1,
            this.названиеDataGridViewTextBoxColumn1,
            this.описаниеDataGridViewTextBoxColumn1,
            this.idпространстваDataGridViewTextBoxColumn});
            this.dataGridView2.DataSource = this.classesBindingSource;
            this.dataGridView2.Location = new System.Drawing.Point(12, 214);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.Size = new System.Drawing.Size(757, 188);
            this.dataGridView2.TabIndex = 1;
            // 
            // dataGridView3
            // 
            this.dataGridView3.AutoGenerateColumns = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.кодDataGridViewTextBoxColumn2,
            this.пространствоDataGridViewTextBoxColumn,
            this.описаниеDataGridViewTextBoxColumn2});
            this.dataGridView3.DataSource = this.spacesBindingSource;
            this.dataGridView3.Location = new System.Drawing.Point(12, 408);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.RowHeadersWidth = 51;
            this.dataGridView3.RowTemplate.Height = 24;
            this.dataGridView3.Size = new System.Drawing.Size(757, 197);
            this.dataGridView3.TabIndex = 2;
            // 
            // databaseDataSet1
            // 
            this.databaseDataSet1.DataSetName = "DatabaseDataSet1";
            this.databaseDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // methodsBindingSource
            // 
            this.methodsBindingSource.DataMember = "methods";
            this.methodsBindingSource.DataSource = this.databaseDataSet1;
            // 
            // methodsTableAdapter
            // 
            this.methodsTableAdapter.ClearBeforeFill = true;
            // 
            // databaseDataSet2
            // 
            this.databaseDataSet2.DataSetName = "DatabaseDataSet2";
            this.databaseDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // classesBindingSource
            // 
            this.classesBindingSource.DataMember = "classes";
            this.classesBindingSource.DataSource = this.databaseDataSet2;
            // 
            // classesTableAdapter
            // 
            this.classesTableAdapter.ClearBeforeFill = true;
            // 
            // databaseDataSet3
            // 
            this.databaseDataSet3.DataSetName = "DatabaseDataSet3";
            this.databaseDataSet3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // spacesBindingSource
            // 
            this.spacesBindingSource.DataMember = "spaces";
            this.spacesBindingSource.DataSource = this.databaseDataSet3;
            // 
            // spacesTableAdapter
            // 
            this.spacesTableAdapter.ClearBeforeFill = true;
            // 
            // кодDataGridViewTextBoxColumn2
            // 
            this.кодDataGridViewTextBoxColumn2.DataPropertyName = "Код";
            this.кодDataGridViewTextBoxColumn2.HeaderText = "Код";
            this.кодDataGridViewTextBoxColumn2.MinimumWidth = 6;
            this.кодDataGridViewTextBoxColumn2.Name = "кодDataGridViewTextBoxColumn2";
            this.кодDataGridViewTextBoxColumn2.Width = 125;
            // 
            // пространствоDataGridViewTextBoxColumn
            // 
            this.пространствоDataGridViewTextBoxColumn.DataPropertyName = "Пространство";
            this.пространствоDataGridViewTextBoxColumn.HeaderText = "Пространство";
            this.пространствоDataGridViewTextBoxColumn.MinimumWidth = 6;
            this.пространствоDataGridViewTextBoxColumn.Name = "пространствоDataGridViewTextBoxColumn";
            this.пространствоDataGridViewTextBoxColumn.Width = 125;
            // 
            // описаниеDataGridViewTextBoxColumn2
            // 
            this.описаниеDataGridViewTextBoxColumn2.DataPropertyName = "Описание";
            this.описаниеDataGridViewTextBoxColumn2.HeaderText = "Описание";
            this.описаниеDataGridViewTextBoxColumn2.MinimumWidth = 6;
            this.описаниеDataGridViewTextBoxColumn2.Name = "описаниеDataGridViewTextBoxColumn2";
            this.описаниеDataGridViewTextBoxColumn2.Width = 125;
            // 
            // кодDataGridViewTextBoxColumn
            // 
            this.кодDataGridViewTextBoxColumn.DataPropertyName = "Код";
            this.кодDataGridViewTextBoxColumn.HeaderText = "Код";
            this.кодDataGridViewTextBoxColumn.MinimumWidth = 6;
            this.кодDataGridViewTextBoxColumn.Name = "кодDataGridViewTextBoxColumn";
            this.кодDataGridViewTextBoxColumn.Width = 125;
            // 
            // названиеDataGridViewTextBoxColumn
            // 
            this.названиеDataGridViewTextBoxColumn.DataPropertyName = "Название";
            this.названиеDataGridViewTextBoxColumn.HeaderText = "Название";
            this.названиеDataGridViewTextBoxColumn.MinimumWidth = 6;
            this.названиеDataGridViewTextBoxColumn.Name = "названиеDataGridViewTextBoxColumn";
            this.названиеDataGridViewTextBoxColumn.Width = 125;
            // 
            // описаниеDataGridViewTextBoxColumn
            // 
            this.описаниеDataGridViewTextBoxColumn.DataPropertyName = "Описание";
            this.описаниеDataGridViewTextBoxColumn.HeaderText = "Описание";
            this.описаниеDataGridViewTextBoxColumn.MinimumWidth = 6;
            this.описаниеDataGridViewTextBoxColumn.Name = "описаниеDataGridViewTextBoxColumn";
            this.описаниеDataGridViewTextBoxColumn.Width = 125;
            // 
            // idклассаDataGridViewTextBoxColumn
            // 
            this.idклассаDataGridViewTextBoxColumn.DataPropertyName = "Id_класса";
            this.idклассаDataGridViewTextBoxColumn.DataSource = this.classesBindingSource;
            this.idклассаDataGridViewTextBoxColumn.DisplayMember = "Название";
            this.idклассаDataGridViewTextBoxColumn.HeaderText = "Класс";
            this.idклассаDataGridViewTextBoxColumn.MinimumWidth = 6;
            this.idклассаDataGridViewTextBoxColumn.Name = "idклассаDataGridViewTextBoxColumn";
            this.idклассаDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.idклассаDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.idклассаDataGridViewTextBoxColumn.ValueMember = "Код";
            this.idклассаDataGridViewTextBoxColumn.Width = 125;
            // 
            // кодDataGridViewTextBoxColumn1
            // 
            this.кодDataGridViewTextBoxColumn1.DataPropertyName = "Код";
            this.кодDataGridViewTextBoxColumn1.HeaderText = "Код";
            this.кодDataGridViewTextBoxColumn1.MinimumWidth = 6;
            this.кодDataGridViewTextBoxColumn1.Name = "кодDataGridViewTextBoxColumn1";
            this.кодDataGridViewTextBoxColumn1.Width = 125;
            // 
            // названиеDataGridViewTextBoxColumn1
            // 
            this.названиеDataGridViewTextBoxColumn1.DataPropertyName = "Название";
            this.названиеDataGridViewTextBoxColumn1.HeaderText = "Название";
            this.названиеDataGridViewTextBoxColumn1.MinimumWidth = 6;
            this.названиеDataGridViewTextBoxColumn1.Name = "названиеDataGridViewTextBoxColumn1";
            this.названиеDataGridViewTextBoxColumn1.Width = 125;
            // 
            // описаниеDataGridViewTextBoxColumn1
            // 
            this.описаниеDataGridViewTextBoxColumn1.DataPropertyName = "Описание";
            this.описаниеDataGridViewTextBoxColumn1.HeaderText = "Описание";
            this.описаниеDataGridViewTextBoxColumn1.MinimumWidth = 6;
            this.описаниеDataGridViewTextBoxColumn1.Name = "описаниеDataGridViewTextBoxColumn1";
            this.описаниеDataGridViewTextBoxColumn1.Width = 125;
            // 
            // idпространстваDataGridViewTextBoxColumn
            // 
            this.idпространстваDataGridViewTextBoxColumn.DataPropertyName = "Id_пространства";
            this.idпространстваDataGridViewTextBoxColumn.DataSource = this.spacesBindingSource;
            this.idпространстваDataGridViewTextBoxColumn.DisplayMember = "Пространство";
            this.idпространстваDataGridViewTextBoxColumn.HeaderText = "Пространство";
            this.idпространстваDataGridViewTextBoxColumn.MinimumWidth = 6;
            this.idпространстваDataGridViewTextBoxColumn.Name = "idпространстваDataGridViewTextBoxColumn";
            this.idпространстваDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.idпространстваDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.idпространстваDataGridViewTextBoxColumn.ValueMember = "Код";
            this.idпространстваDataGridViewTextBoxColumn.Width = 125;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1066, 647);
            this.Controls.Add(this.dataGridView3);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.methodsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.classesBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.databaseDataSet3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.spacesBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private DatabaseDataSet databaseDataSet;
        private System.Data.OleDb.OleDbConnection oleDbConnection1;
        private System.Data.OleDb.OleDbCommand oleDbSelectCommand1;
        private System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter1;
        private System.Data.OleDb.OleDbCommand oleDbSelectCommand2;
        private System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter2;
        private System.Data.OleDb.OleDbCommand oleDbSelectCommand3;
        private System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter3;
        private System.Data.OleDb.OleDbCommand oleDbDeleteCommand;
        private System.Data.OleDb.OleDbCommand oleDbInsertCommand;
        private System.Data.OleDb.OleDbCommand oleDbUpdateCommand;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridView dataGridView3;
        private DatabaseDataSet1 databaseDataSet1;
        private System.Windows.Forms.BindingSource methodsBindingSource;
        private DatabaseDataSet1TableAdapters.methodsTableAdapter methodsTableAdapter;
        private DatabaseDataSet2 databaseDataSet2;
        private System.Windows.Forms.BindingSource classesBindingSource;
        private DatabaseDataSet2TableAdapters.classesTableAdapter classesTableAdapter;
        private DatabaseDataSet3 databaseDataSet3;
        private System.Windows.Forms.BindingSource spacesBindingSource;
        private DatabaseDataSet3TableAdapters.spacesTableAdapter spacesTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn кодDataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn пространствоDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn описаниеDataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn кодDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn названиеDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn описаниеDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewComboBoxColumn idклассаDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn кодDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn названиеDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn описаниеDataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewComboBoxColumn idпространстваDataGridViewTextBoxColumn;
    }
}

