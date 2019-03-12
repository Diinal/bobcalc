using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bobcalc
{
    public partial class Page3 : Form
    {
        public Page3()
        {
            InitializeComponent();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
            excel.excelapp.Windows[1].Close(false);
        }

        private void MuteButton_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private System.Drawing.Point MouseHook;

        private void Top_Label_MouseMove(object sender, MouseEventArgs e)
        {
            Cursor = Cursors.Default;
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            else
            {
                Location = new System.Drawing.Point((Size)Location - (Size)MouseHook + (Size)e.Location);
                Cursor = Cursors.Hand;
                Top_Label.Focus();
            }
        }

        private void CloseButton_MouseHover(object sender, EventArgs e)
        {
            CloseButton.BackColor = Color.DarkGray;
        }

        private void CloseButton_MouseLeave(object sender, EventArgs e)
        {
            CloseButton.BackColor = Color.AliceBlue;
        }

        private void MuteButton_MouseHover(object sender, EventArgs e)
        {
            MuteButton.BackColor = Color.DarkGray;
        }

        private void MuteButton_MouseLeave(object sender, EventArgs e)
        {
            MuteButton.BackColor = Color.AliceBlue;
        }

        private void Back_Arrow_Button_MouseHover(object sender, EventArgs e)
        {
            Back_Arrow_Button.BackColor = Color.DarkGray;
        }

        private void Back_Arrow_Button_MouseLeave(object sender, EventArgs e)
        {
            Back_Arrow_Button.BackColor = Color.AliceBlue;
        }

        private void Back_Arrow_Button_Click(object sender, EventArgs e)
        {
            Page2 pag2 = new Page2();
            pag2.Show();
            this.Hide();
        }
    }
}
