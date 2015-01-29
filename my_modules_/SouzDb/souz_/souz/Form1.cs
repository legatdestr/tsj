using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

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
            FSettings fSettings = new FSettings();
            fSettings.ShowDialog();
            }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
            {
            AboutBox1 fAbout = new AboutBox1();
            fAbout.ShowDialog(this);
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
            // TODO: данная строка кода позволяет загрузить данные в таблицу "archiveDataSet.selLS". При необходимости она может быть перемещена или удалена.
            this.selLSTableAdapter.Fill(this.archiveDataSet.selLS);
            showArchiveManager();
            ConnectFormToArchive();

            }


        // строит условную часть запроса по данным формы
        private string buildFormSQLQuery()
            {
            string sql =
                "SELECT FIO, DOM, DOML, DOMP, KV, KVL, LS, NZEU, UL, spUL_NAIM, spZEU_NAIM, PEN_S1, ITGNW, SAL1W, SAL2W FROM selLS ";

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
                    sql += " ORDER BY spUL_NAIM ASC ";
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
            //if (archiveCon.State == System.Data.ConnectionState.Closed) { archiveCon.Open() ; }

            archiveCon = this.selLSTableAdapter.Connection;
            string sql = buildFormSQLQuery();
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, archiveCon);
            this.archiveDataSet.selLS.Clear();
            richTextBox1.AppendText(buildFormSQLQuery() + "\r\n");
            adapter.Fill(this.archiveDataSet.selLS); 
            //this.selLSTableAdapter.Fill(this.archiveDataSet.selLS);
            
           // MessageBox.Show(buildFormSQLQuery());
            this.Cursor = Cursors.Default;
        }

    //            DateTime dt = new DateTime(2013, 2, 1);
    //            db.AccessImport.DbImport.DoImport(dt, @"C:\KVPL\", openFileDialog1.FileName);
                

     
    }
}
