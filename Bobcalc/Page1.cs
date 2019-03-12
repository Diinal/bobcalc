using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Bobcalc
{
    public partial class Page1 : Form
    {
        Page2 pag2 = new Page2();
        public Page1()
        {
            InitializeComponent();
            this.DoubleBuffered = true;
        }

        private void Start_Button_Click(object sender, EventArgs e)
        {
            //Подгрузка второй формы
            pag2.Show();
            this.Hide();
            
        }

        private void CloseButton_Click_1(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void MuteButton_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            pag2.WindowState = FormWindowState.Minimized;
        }
            
        private System.Drawing.Point MouseHook;

        //Анимации наведения на кнопки управления

        private void MuteButton_MouseHover_1(object sender, EventArgs e)
        {
        }

        private void MuteButton_MouseLeave_1(object sender, EventArgs e)
        {
            MuteButton.BackColor = Color.AliceBlue;
        }

        private void CloseButton_MouseHover_1(object sender, EventArgs e)
        {
            CloseButton.BackColor = Color.DarkGray;
        }

        private void CloseButton_MouseLeave_1(object sender, EventArgs e)
        {
            CloseButton.BackColor = Color.AliceBlue;
        }
        //перетаскивание окна за лейбл
        private void Top_Label_MouseMove_1(object sender, MouseEventArgs e)
        {
            Cursor = Cursors.Default;
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            else
            {
                Location = new System.Drawing.Point((Size)Location - (Size)MouseHook + (Size)e.Location);
                pag2.Location = new System.Drawing.Point((Size)Location - (Size)MouseHook + (Size)e.Location);
                Cursor = Cursors.Hand;
            }
        }

        private void путьДляExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string value = "";
            if (Program.InputBox("Путь к Excel документу", "Введите путь к документу Excel:", ref value) == DialogResult.OK)
            {
                string path = value;
                Microsoft.Win32.RegistryKey excelpath = null;
                try
                {
                    excelpath = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                    if (excelpath != null)
                        excelpath.SetValue("ExcelPath", path);
                }
                finally
                {
                    if (excelpath != null) excelpath.Close();
                }
                MessageBox.Show("Пожалуйста перезапустите приложение");
            }
        }

        private void путьДляPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string value = "";
            if (Program.InputBox("Путь для PDF документа", "Введите путь для коммерческкого предложения:", ref value) == DialogResult.OK)
            {
                string path = value;
                Microsoft.Win32.RegistryKey pdfpath = null;
                try
                {
                    pdfpath = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                    if (pdfpath != null)
                        pdfpath.SetValue("PDFPath", path);
                }
                finally
                {
                    if (pdfpath != null) pdfpath.Close();
                }

            }
        }
    }
}
