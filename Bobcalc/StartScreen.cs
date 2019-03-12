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
    public partial class StartScreen : Form
    {
        public StartScreen()
        {
            InitializeComponent();
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {   
            base.Opacity += 0.1;

            if (base.Opacity == 1.0)
            {
                timer1.Stop();
            }
        }
        
    }
}
