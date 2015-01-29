using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;
namespace souz
{
    public partial class FMain : Form
    {
        public FMain()
        {
            InitializeComponent();
        }

        // поле, хранящее объект connection к Архиву
        public OleDbConnection archiveCon = new OleDbConnection();

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
            {
                using (FSettings fSettings = new FSettings())
                {
                    fSettings.ShowDialog();
                }
            }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
            {
                using (AboutBox1 fAbout = new AboutBox1())
                {
                    fAbout.ShowDialog(this);
                }
            }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
            {
            showArchiveManager();
            }

        // показываем форму управления архивом
        private void showArchiveManager()
            {
            FArchiveManager fArchiveMan = new FArchiveManager();
            fArchiveMan.ShowDialog(this);
            addSpravkaToForm("Работа с менеджером архива завершена. Можно приступать к поиску должников.");
            }

        // показываем данные загруженного месяца в заголовке
        private void showLoadedMonth()
            {
            if (Properties.Settings.Default.archiveUpdateDt == null) { return; }
            DateTime dt = Properties.Settings.Default.archiveUpdateDt;
            string m = dt.Month.ToString();
            if (m.Length < 2) { m = "0" + m; }
            this.Text = Properties.Settings.Default.mainFormCaption + " - " + m + "." + dt.Year;
            }

        // Обработка события загрузки формы
        private void FMain_Load(object sender, EventArgs e)
            {
            showArchiveManager();
            // Подгружает справочники, инициализирует комбобоксы и т.д.
            ConnectFormToArchive();
            // Таймер запускает очищение панели статуса, если в ней засиделось сообщение
            this.timer1.Enabled = true;

            }


        // строит условную часть запроса по данным формы
        private string buildFormSQLQuery(bool whereOnly = false)
            {
                string sql = "";
            if (whereOnly == false)
            {
                // если выбран поиск должников как в KVPL
                if (checkBox2.Checked)
                {
                    sql =
                    "SELECT FIO, UL, DOM, DOML, DOMP, KV, KVL, LS, NZEU, LSX, SAL1W, ITGNW, PEN, SAL2W, COMMENT, INS_DATE, OPLATA_DO, DOLG_POSLE_OPLATY, Street,Company FROM selLSKakVKVPL ";
                }
                else
                {
                    sql =
                        "SELECT FIO, UL, DOM, DOML, DOMP, KV, KVL, LS, NZEU, LSX, SAL1W, ITGNW, PEN, SAL2W, COMMENT, INS_DATE, OPLATA_DO, DOLG_POSLE_OPLATY, Street,Company FROM selLS ";
                }
            }
            string where = null;
            //жилое           
            if (radioButton1.Checked)
                {
                where = " (TIP <> 5)";
                }
            // нежилое
            if (radioButton2.Checked)
                {
                where = " (TIP = 5)";
                }
            // искать только в выбранной базе
            if (radioButton5.Checked)
                {
                if (where != null) { where += " AND "; }
                where += " (NZEU = " + comboBox1.SelectedValue.ToString() + ")";
                }
            // искать по номеру лицевого счета
            if (maskedTextBox1.Text != "")
                {
                if (where != null) { where += " AND "; }
                where += " (LS=" + maskedTextBox1.Text + ")";
                }
            // по фио
            if (textBox1.Text.Length > 1)
                {
                if (where != null) { where += " AND "; }
                where += @" (FIO like '%" + textBox1.Text + @"%')";
                }
            // улица
            if (radioButton4.Checked)
                {
                if (where != null) { where += " AND "; }
                where += @" (UL = " + comboBox2.SelectedValue.ToString() + ")";
                }
            // дом
            if (numericUpDown1.Value > 0)
                {
                if (where != null) { where += " AND "; }
                where += @" ( DOM = " + numericUpDown1.Value.ToString() + ")";
                }

            // корпус
            if (textBox3.Text.Length > 0)
                {
                if (where != null) { where += " AND "; }
                where += @" ( DOML = '" + textBox3.Text + "')";
                }
            // долг от
            if (numericUpDown2.Value > 0)
                {
                if (where != null) { where += " AND "; }
                where += @"(ITGNW >= " + numericUpDown2.Value.ToString() + ")";
                }
            // долг до
            if (numericUpDown3.Value > 0)
                {
                if (where != null) { where += " AND "; }
                where += @"(ITGNW <= " + numericUpDown3.Value.ToString() + ")";
                }
            // сальдо н от
            if (numericUpDown5.Value > 0)
                {
                if (where != null) { where += " AND "; }
                where += @" (SAL1W >= " + numericUpDown5.Value.ToString() + ") ";
                }
            // сальдо к от
            if (numericUpDown4.Value > 0)
                {
                if (where != null) { where += " AND "; }
                where += @" (SAL2W >= " + numericUpDown4.Value.ToString() + ") ";
                }

            // помечен комментарием и на сегодня не оплатил
            if (checkBox1.Checked)
            {
                if (where != null) { where += " AND "; }
                where += @" ( (OPLATA_DO <=  Now()) AND (SAL2W >= DOLG_POSLE_OPLATY)"  + ") ";
            }

            // условия вообще есть? если есть - добавляем оператор where.
            if (where != null) { sql += " WHERE " + where; }

            // сортировка:
            if (comboBox3.Text == "НачислИтг(Долг)")
                {
                sql += " ORDER BY ITGNW DESC ";
                }
            else
                if (comboBox3.Text == "ФИО")
                    {
                    sql += " ORDER BY FIO ASC ";
                    }
                else
                    if (comboBox3.Text == "Адрес (Улица и номер дома)")
                        {
                        sql += " ORDER BY UL, DOM ASC ";
                        }




            return sql;
            }


        // загружаем данные из справочников в выпадающие списки
        public void ConnectFormToArchive()
            {
            // подключаем список компаний
            if (archiveCon.State == System.Data.ConnectionState.Closed)
                {
                try
                    {
                    archiveCon.ConnectionString = Properties.Settings.Default.archiveConnectionString;
                    archiveCon.Open() ;
                    }
                catch (OleDbException e)
                    {
                    MessageBox.Show("Не удалось подключиться к архиву. " + e.Message , "Ошибка работы с архивом!");
                    }
                }
            OleDbDataAdapter spZeuAdapter = new OleDbDataAdapter("SELECT spZEU.ZEU, spZEU.NAIM FROM spZEU order by spZEU.NAIM", archiveCon);
            DataSet DbAccessDataSet = new DataSet();
            spZeuAdapter.Fill(DbAccessDataSet, "spZeu");

            if ((DbAccessDataSet.Tables.Count > 0) && (DbAccessDataSet.Tables[0].Rows.Count > 0))
                {
                comboBox1.DataSource = DbAccessDataSet.Tables[0];
                comboBox1.DisplayMember = "NAIM";
                comboBox1.ValueMember = "ZEU";
                }

            // подключаем список улиц
            OleDbDataAdapter spUlAdapter = new OleDbDataAdapter("SELECT spUL.UL, spUL.NAIM FROM spUL order by spUL.NAIM", archiveCon);
            spUlAdapter.Fill(DbAccessDataSet, "spUL");

            if ((DbAccessDataSet.Tables.Count > 0) && (DbAccessDataSet.Tables[1].Rows.Count > 0))
                {
                comboBox2.DataSource = DbAccessDataSet.Tables[1];
                comboBox2.DisplayMember = "NAIM";
                comboBox2.ValueMember = "UL";
                }
            addSpravkaToForm("Подключение справочников завершено!", true);
            }

        private void начатьРаботуToolStripMenuItem_Click(object sender, EventArgs e)
            {
            ConnectFormToArchive();
            }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
            {
            comboBox2.Enabled = radioButton4.Checked;
            }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
            {
            comboBox1.Enabled = radioButton5.Checked;
            }

        // Обработка кнопки поиска
        private void button1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            // открываем connection
            //archiveCon.ConnectionString = Properties.Settings.Default.archiveConnectionString;
            
            archiveCon = this.selLSTableAdapter.Connection;
            if (archiveCon.State == System.Data.ConnectionState.Closed) { archiveCon.Open(); }

            string sql = buildFormSQLQuery();
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, archiveCon);
            this.archiveDataSet.selLS.Clear();
            richTextBox1.AppendText(buildFormSQLQuery() + "\r\n");
            adapter.Fill(this.archiveDataSet.selLS); 
            //this.selLSTableAdapter.Fill(this.archiveDataSet.selLS);
            
           // MessageBox.Show(buildFormSQLQuery());
            this.Cursor = Cursors.Default;
        }

        // обработка кнопки удалить комментарий
        private void button3_Click(object sender, EventArgs e)
            {

            if ( archiveDataSet.selLS.Count > 0)
                {
                if ((cOMMENTRichTextBox.Text != "")||(oPLATA_DOTextBox.Text != ""))
                    {
                    if (MessageBox.Show("Вы подтверждаете, что хотите удалить комментарий?", "Запрос на удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                        string deleteSql = "DELETE FROM COMMENTS WHERE COMMENTS.LS = "
                        + (selLSBindingSource.Current as DataRowView)["LS"].ToString();
                        //MessageBox.Show(deleteSql);
                        OleDbCommand del = new OleDbCommand(deleteSql, this.archiveCon);
                        del.ExecuteNonQuery();
                        cOMMENTRichTextBox.Text = null;
                        oPLATA_DOTextBox.Text = null;
                        addSpravkaToForm("Удален комментарий для лицевого счёта: " + 
                            (selLSBindingSource.Current as DataRowView)["LS"].ToString());
                        }
                    }
                
                }
            }

        // добавление информации в панель статуса
        public void addSpravkaToForm(string st = "", bool add = false)
            {
            this.timer1.Enabled = false;
            this.timer1.Enabled = true;
            if (add)
                {
                this.toolStripStatusLabel1.Text += ' ' + st;
                return;
                }
            this.toolStripStatusLabel1.Text = st;
            
            }

        // почему-то не работает update
        private void button2_Click(object sender, EventArgs e)
            {

            if (archiveDataSet.selLS.Count > 0)
                {
                string sql;
                if ((richTextBox2.Text == "") && (textBox2.Text == "") && (maskedTextBox2.Text == ""))
                { return; }
                if (textBox2.Text == "")
                    {
                    MessageBox.Show("Заполните поле Оплатить до.");
                    return;
                    }
                if (maskedTextBox2.Text == "")
                    {
                    MessageBox.Show("Заполните поле \"Долг после оплаты не более.\"");
                    return;
                    }
                try
                    {
                    DateTime dt = DateTime.Parse(textBox2.Text);
                    }
                catch (FormatException ex)
                    {
                    MessageBox.Show("Заполните поле \"Оплатить до.\" " + ex.Message);
                    return;
                    }

                

                // если необходим update
                if ((cOMMENTRichTextBox.Text != "") || (oPLATA_DOTextBox.Text != ""))
                    {
                   
                    if (archiveCon.State == System.Data.ConnectionState.Closed) { archiveCon.Open(); }

                    //OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.archiveConnectionString1);
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = archiveCon;
                    cmd.CommandText = "UPDATE Comments "
                        + "SET COMMENT = @COM, OPLATA_DO = @DOOPL, DOLG_POSLE_OPLATY = @DOLG_POS "
                        + " WHERE (LS = @L) AND (NZEU = @NZ) ";
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@L", OleDbType.Integer).Value = (selLSBindingSource.Current as DataRowView)["LS"];
                    cmd.Parameters.Add("@NZ", OleDbType.Integer).Value = (selLSBindingSource.Current as DataRowView)["NZEU"];
                    cmd.Parameters.Add("@COM", OleDbType.VarChar).Value = richTextBox2.Text;
                    cmd.Parameters.Add("@DOOPL", OleDbType.DBTimeStamp).Value = DateTime.Parse(textBox2.Text);
                    cmd.Parameters.Add("@DOLG_POS", OleDbType.Numeric).Value = maskedTextBox2.Text;


                    sql =
                        "UPDATE COMMENTS "
                    + " SET COMMENT = '" + richTextBox2.Text + "'"
                    + ", OPLATA_DO = '" + DateTime.Parse(textBox2.Text) + "'"
                    + ", DOLG_POSLE_OPLATY = " + maskedTextBox2.Text
                    + " WHERE LS = " + (selLSBindingSource.Current as DataRowView)["LS"].ToString()
                    + " AND NZEU = " + (selLSBindingSource.Current as DataRowView)["NZEU"].ToString()
                    ;
                    cmd.Parameters.Clear();
                    cmd.CommandText = sql;

                    try
                        {
                        cmd.ExecuteNonQuery();
                        }
                    catch (Exception ex)
                        {
                        MessageBox.Show("Невозможно выполнить обновление комментария! \r\n Подробности: " + ex.Message);
                        }

                    richTextBox1.Text += "\r\n" + cmd.CommandText;
                    (selLSBindingSource.Current as DataRowView).BeginEdit();
                    (selLSBindingSource.Current as DataRowView)["COMMENT"] = richTextBox2.Text;
                    (selLSBindingSource.Current as DataRowView)["OPLATA_DO"] = dateTimePicker1.Value;
                    (selLSBindingSource.Current as DataRowView)["INS_DATE"] = DateTime.Today;
                    (selLSBindingSource.Current as DataRowView).EndEdit();
                    oPLATA_DOTextBox.Text = textBox2.Text;
                    addSpravkaToForm( "Обновлена информация для лицевого счета: " + (selLSBindingSource.Current as DataRowView)["LS"]);
                    }
                else // значит Insert
                    {
                    if (archiveCon.State == System.Data.ConnectionState.Closed) { archiveCon.Open(); }

                    //OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.archiveConnectionString1);
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = archiveCon;
                    cmd.CommandText = "INSERT INTO Comments (LS, NZEU, COMMENT, OPLATA_DO, DOLG_POSLE_OPLATY) "
                        + "VALUES (@LS, @NZEU, @COMMENT, @OPLATA_DO, @DOLG_POSLE_OPLATY)";
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@LS", OleDbType.Integer).Value = (selLSBindingSource.Current as DataRowView)["LS"];
                    cmd.Parameters.Add("@NZEU", OleDbType.Integer).Value = (selLSBindingSource.Current as DataRowView)["NZEU"];
                    cmd.Parameters.Add("@COMMENT", OleDbType.VarChar).Value = richTextBox2.Text;
                    cmd.Parameters.Add("@OPLATA_DO", OleDbType.DBTimeStamp).Value = DateTime.Parse(textBox2.Text);
                    cmd.Parameters.Add("@DOLG_POSLE_OPLATY", OleDbType.Numeric).Value = maskedTextBox2.Text;

                    cmd.ExecuteNonQuery();

                    richTextBox1.Text += "\r\n" + cmd.CommandText;
                    (selLSBindingSource.Current as DataRowView).BeginEdit();
                    (selLSBindingSource.Current as DataRowView)["COMMENT"] = richTextBox2.Text;
                    (selLSBindingSource.Current as DataRowView)["OPLATA_DO"] = dateTimePicker1.Value;
                    (selLSBindingSource.Current as DataRowView)["INS_DATE"] = DateTime.Today;
                    (selLSBindingSource.Current as DataRowView).EndEdit();
                    oPLATA_DOTextBox.Text = textBox2.Text;
                    addSpravkaToForm("Добавлен комментарий для лицевого счета: " + (selLSBindingSource.Current as DataRowView)["LS"]);
                    }
                                


    

/*
                OleDbConnection myOleDbConnection = new OleDbConnection(Properties.Settings.Default.archiveConnectionString1);
                OleDbCommand myOleDbCommand = myOleDbConnection.CreateCommand();
                myOleDbCommand.CommandText = sql;
                myOleDbConnection.Open();
               // myOleDbCommand.ExecuteNonQuery();
                richTextBox1.Text += "\r\n"+ sql;
                myOleDbConnection.Close();
                //sqlCom.ExecuteNonQuery();
                cOMMENTRichTextBox.Text = richTextBox2.Text;
                oPLATA_DOTextBox.Text = maskedTextBox2.Text;
                MessageBox.Show("Готово!");
                //con.Close();
*/
                }
            }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
            {
            textBox2.Text = dateTimePicker1.Value.ToString("d");
            }



        private void numericUpDown2_Validated(object sender, EventArgs e)
            {
            if (((sender as NumericUpDown).Text == "") || ((sender as NumericUpDown).Text == null))
                {
                (sender as NumericUpDown).Text = "0";
                }
            }

        private void exportToWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataExport.WordExport.Export(archiveDataSet.Tables[0], @"D:\l\WORK2\my_soft\my_modules\SouzDb\souz\souz\templates\sud2.doc", @"D:\l\WORK2\my_soft\my_modules\SouzDb\souz\souz\templates\res.doc");
        }

        private void doExportToWord()
        {
            // Если вообще есть информация
            if (archiveDataSet.selLS.Count > 0)
            {
                // Открываем соединение с архивом
                if (archiveCon.State == System.Data.ConnectionState.Closed) { archiveCon.Open(); }

                // Для начала удаляем старые квитанции
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = archiveCon;
                string sql = "DELETE FROM Kvit";
                cmd.CommandText = sql;
                try
                    {
                    cmd.ExecuteNonQuery();
                    }
                catch (Exception ex)
                    {
                    MessageBox.Show("Возникла ошибка при удалении старых квитанций. " + ex.Message);
                    addSpravkaToForm("Возникла ошибка при удалении старых квитанций. " + ex.Message);                    
                    return;
                    }
            // Запрос на заполнение таблицы с квитанциями. Важно чтобы условия отбора были такие же как и в фильтре формы.
                sql = 
                     "INSERT INTO "
                        + "Kvit( "
                        + "LS, "  
                        + "ST_FIO, " 
                        + "DT_SOST_NA, "  
                        + "NUM_SUMMA, "  
                        + "DT_OPLATA_DO, "  
                        + "NUM_OPLATA_USL, "  
                        + "NUM_PEN, "  
                        + "NUM_GOS_POSH, "  
                        + "NUM_ITOGO, "  
                        + "ST_ULICA, "  
						+ "ST_DOML, " 
                        + "NUM_DOM, "  
                        + "ST_DOMP, "  
                        + "ST_KV, "
						+ "ST_KVL, "
                        + "NUM_NZEU"
                        + ") "
                    + "SELECT "
                        + "LS, "  
                        + "FIO, "  
                        + "DATE(), " 
                        + " Round(SAL2W,2), " 
                        + " OPLATA_DO, "  
                        + " @NUM_OPLATA_USL, "
                        + " Round(SAL2W/10, 2) as PEN, " 
                        + " @NUM_GOS_POSH, "
                        + " ROUND(SAL2W+@NUM_OPLATA_USL+SAL2W/10+@NUM_GOS_POSH,2) as NUM_ITOGO, " 
                        + " Street, " 
                        + " DOML, " 
                        + " DOM, " 
						+ " DOMP, " 
                        + " KV , "
						+ " KVL, "
                        + " NZEU"
                    + " FROM selLS "
                    + buildFormSQLQuery(true);
                // если отмечен флажок как в kvpl - :
                if (checkBox2.Checked)
                    {
                    sql =
                         "INSERT INTO "
                            + "Kvit( "
                            + "LS, "
                            + "ST_FIO, "
                            + "DT_SOST_NA, "
                            + "NUM_SUMMA, "
                            + "DT_OPLATA_DO, "
                            + "NUM_OPLATA_USL, "
                            + "NUM_PEN, "
                            + "NUM_GOS_POSH, "
                            + "NUM_ITOGO, "
                            + "ST_ULICA, "
                            + "ST_DOML, "
                            + "NUM_DOM, "
                            + "ST_DOMP, "
                            + "ST_KV, "
                            + "ST_KVL, "
                            + "NUM_NZEU"
                            + ") "
                        + "SELECT "
                            + "LS, "
                            + "FIO, "
                            + "DATE(), "
                            + " Round(SAL2W,2), "
                            + " OPLATA_DO, "
                            + " @NUM_OPLATA_USL, "
                            + " Round(SAL2W/10, 2) as PEN, "
                            + " @NUM_GOS_POSH, "
                            + " ROUND(SAL2W+@NUM_OPLATA_USL+SAL2W/10+@NUM_GOS_POSH,2) as NUM_ITOGO, "
                            + " Street, "
                            + " DOML, "
                            + " DOM, "
                            + " DOMP, "
                            + " KV , "
                            + " KVL, "
                            + " NZEU "
                        + " FROM selLSKakVKVPL "
                        + buildFormSQLQuery(true);
                    }


                cmd.Parameters.Add("@NUM_OPLATA_USL", OleDbType.Double).Value = Properties.Settings.Default.oplataUsl;
                cmd.Parameters.Add("@NUM_GOS_POSH", OleDbType.Double).Value = Properties.Settings.Default.gosPoshlina;
                cmd.CommandText = sql;
                try
                {
                    cmd.ExecuteNonQuery();
                    sql = "UPDATE Kvit SET DT_OPLATA_DO = DateAdd(\"ww\",2,Date()) WHERE DT_OPLATA_DO IS NULL";
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                catch (Exception e)
                { 
                    MessageBox.Show(e.Message); 
                    
                }
                finally
                {
                if ((Properties.Settings.Default.pathToKvitTemplate != null)
                    && System.IO.File.Exists(Properties.Settings.Default.pathToKvitTemplate)
                    && (Properties.Settings.Default.pathToKvitTemplate != ""))
                    {
                    try
                        {
                        Process.Start(Properties.Settings.Default.pathToKvitTemplate);
                        }
                    catch (Win32Exception ex)
                        {
                            MessageBox.Show("An error occurred when opening the associated file.");
                            addSpravkaToForm("Экспорт данных не был выполнен.");
                        }
                    }
                else
                    {
                    MessageBox.Show("Необходимо в настройках указать путь к файлу с шаблоном Word.");
                    addSpravkaToForm("Экспорт данных не был выполнен.");
                    using (FSettings fSettings = new FSettings())
                        {
                        fSettings.ShowDialog();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет данных для экспорта! ");
            }
        }

        private void exportToWord2ToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void экспортВКвитанцииToolStripMenuItem_Click(object sender, EventArgs e)
        {

            doExportToWord();
        }

        private void экспорт2ToolStripMenuItem_Click(object sender, EventArgs e)
            {
            System.Data.DataTable dt  = archiveDataSet.Tables["selLS"];
            OpenFileDialog opDial = new OpenFileDialog();
            opDial.Title = "Выберите шаблон для формирования квитанций:";
            opDial.Filter = "Ms Word Шаблон|*.doc";
            opDial.DefaultExt = "doc";
            opDial.CheckFileExists = true;
            if (opDial.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                SaveFileDialog saveDial = new SaveFileDialog();
                saveDial.Title = "Выберите, куда сохранить документ";
                saveDial.Filter = "Документ Ms Word |*.doc";
                if (saveDial.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        DataExport.WordExport.Export(dt, opDial.FileName, saveDial.FileName);

                    }
                
                }
            }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            // если отмечен поиск как в KVPL
            bool ch = !(sender as CheckBox).Checked;
            if ((sender as CheckBox).Checked) {
                archiveDataSet.Clear();         
            }
            numericUpDown4.Value = 0;
            numericUpDown5.Value = 0;
            numericUpDown4.Enabled = ch;
            numericUpDown5.Enabled = ch;

            numericUpDown2.Value = 0;
            numericUpDown3.Value = 0;
            numericUpDown2.Enabled = ch;
            numericUpDown3.Enabled = ch;

        }

        private void timer1_Tick(object sender, EventArgs e)
            {
            this.toolStripStatusLabel1.Text = "";
            }

        private void exportToXMLToolStripMenuItem_Click(object sender, EventArgs e)
            {
            if (archiveDataSet.selLS.Count < 1) { return;}
            SaveFileDialog sD = new SaveFileDialog();
            sD.Filter = "XML файлы |*.xml";
            sD.DefaultExt = "xml";
            sD.Title = "Выберите, куда сохранить документ";
            if (sD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                archiveDataSet.WriteXml(sD.FileName);
                addSpravkaToForm("Данные экспортированы: " + sD.FileName + ".  Вы можете открыть файл, например, в Microsoft Excel.");
                }
            }

        private void selLSDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
            {

            }


     
    }
}
