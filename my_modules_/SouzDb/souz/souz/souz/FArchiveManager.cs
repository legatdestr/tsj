using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace souz
    {
    public partial class FArchiveManager : Form
        {
        public FArchiveManager()
            {
            InitializeComponent();
            }

        private void button1_Click(object sender, EventArgs e)
            {
            FSettings fSettings = new FSettings();
            fSettings.ShowDialog();
            }

        private void button2_Click(object sender, EventArgs e)
            {
            toolStripStatusLabel1.Text = "Идет загрузка. Пожалуйста, ждите...";
            if ((Properties.Settings.Default.kvplDbPath == null) || (Properties.Settings.Default.kvplDbPath == ""))
                {
                
                MessageBox.Show("Укажите путь к базе данных!");
                FSettings fSettings = new FSettings();
                fSettings.ShowDialog(this);
                return;
                }
            if (System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(Properties.Settings.Default.kvplDbPath)))
                {
                db.AccessImport.DbImport.DoImport(dateTimePicker1.Value,
                    Properties.Settings.Default.kvplDbPath, 
                    AppDomain.CurrentDomain.BaseDirectory);
                toolStripStatusLabel1.Text = "Загрузка завершена!";
                Properties.Settings.Default.archiveUpdateDt = dateTimePicker1.Value;
                string m = dateTimePicker1.Value.Month.ToString();
                if (m.Length < 2) { m = "0" + m.ToString(); }
                Application.OpenForms[0].Text  =
                    Properties.Settings.Default.mainFormCaption + 
                    " - " + m + "." +
                    dateTimePicker1.Value.Year.ToString();
                MessageBox.Show("Операция поиска данных завершена!");
                Close();                
                }
            else
            { toolStripStatusLabel1.Text = "Данные не загрузились..." + Properties.Settings.Default.kvplDbPath; }
           //Close();
            }

        private void button3_Click(object sender, EventArgs e)
            {
            Close();
            }
        }
    }
