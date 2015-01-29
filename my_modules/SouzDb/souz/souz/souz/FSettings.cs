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
    public partial class FSettings : Form
        {
        public FSettings()
            {
            InitializeComponent();
            }

        private void FSettings_Load(object sender, EventArgs e)
            {
            label1.Dock = DockStyle.Fill;
            label3.Dock = DockStyle.Fill;
            label4.Dock = DockStyle.Fill;
            button1.Dock = DockStyle.Fill;
            button4.Dock = button1.Dock;
            if ((Properties.Settings.Default.kvplDbPath != null) && (Properties.Settings.Default.kvplDbPath != ""))
                {
                button1.Text = Properties.Settings.Default.kvplDbPath;
                }
            if ((Properties.Settings.Default.pathToKvitTemplate != null) && (Properties.Settings.Default.pathToKvitTemplate != ""))
                {
                button4.Text = Properties.Settings.Default.pathToKvitTemplate;
                }
            textBox1.Text = Properties.Settings.Default.oplataUsl.ToString();
            textBox2.Text = Properties.Settings.Default.gosPoshlina.ToString();

            }

        private void button1_Click(object sender, EventArgs e)
            {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                button1.Text = openFileDialog1.FileName;
                }
            }

        private void SaveSettingsFromForm()
            {
            Properties.Settings.Default.kvplDbPath = button1.Text;
            Properties.Settings.Default.oplataUsl = Convert.ToDouble(textBox1.Text);
            Properties.Settings.Default.gosPoshlina = Convert.ToDouble(textBox2.Text);
            Properties.Settings.Default.pathToKvitTemplate = button4.Text;
            Properties.Settings.Default.Save();

            }

        private void button2_Click(object sender, EventArgs e)
            {
            if (System.IO.File.Exists(button1.Text))
                {
                SaveSettingsFromForm();
                (Application.OpenForms["FMain"] as FMain).addSpravkaToForm("Настройки сохранены...");
                this.Close();
                }
            else
                {
                MessageBox.Show("Укажите путь к файлу базы данных: kvpl.mdb.", "Ошибка!", MessageBoxButtons.OK);

                }

            }



        private void button3_Click(object sender, EventArgs e)
            {
            Close();
            }

        private void button4_Click(object sender, EventArgs e)
            {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                button4.Text = openFileDialog2.FileName;
                }
            }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
            {

            }

        }
    }
