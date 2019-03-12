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

    //Страница два
    public partial class MainPage : Form
    {
        Gate gate = new Gate();
        Wicket wicket = new Wicket();
        Auto auto = new Auto();

        // static string path = "" /*"E:\\Bobmaster\\Bobmaster Calculator"*/;

        public MainPage()
        {
            InitializeComponent();
            this.DoubleBuffered = true;
        }

        public int GetPrice()
        {
            try {
                int[,] pricearray = new int[5, 13];
                switch (gate.execution)
                {
                    case "рама без обшивки":
                        for (int i = 0; i < 5; i++)
                            for (int j = 0; j < 13; j++)
                            {   
                                Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[i + 4, j + 4];                               
                                pricearray[i, j] = Convert.ToInt32(cell.Value);
                            }
                        gate.price = pricearray[gate.convertheight(), gate.convertwidth()];                       
                        return pricearray[gate.convertheight(), gate.convertwidth()];
                        break;

                    case "проф.лист C-8 одна сторона":
                        for (int i = 0; i < 5; i++)
                            for (int j = 0; j < 13; j++)
                            {
                                Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[i + 13, j + 4];
                                pricearray[i, j] = Convert.ToInt32(cell.Value);
                            }
                        gate.price = pricearray[gate.convertheight(), gate.convertwidth()];
                        return pricearray[gate.convertheight(), gate.convertwidth()];
                        break;

                    case "проф.лист C-8 две стороны":
                        for (int i = 0; i < 5; i++)
                            for (int j = 0; j < 13; j++)
                            {
                                Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[i + 22, j + 4];
                                pricearray[i, j] = Convert.ToInt32(cell.Value);
                            }
                        gate.price = pricearray[gate.convertheight(), gate.convertwidth()];
                        return pricearray[gate.convertheight(), gate.convertwidth()];
                        break;

                    case "гладкий лист 2 мм":
                        for (int i = 0; i < 5; i++)
                            for (int j = 0; j < 13; j++)
                            {
                                Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[i + 31, j + 4];
                                pricearray[i, j] = Convert.ToInt32(cell.Value);
                            }
                        gate.price = pricearray[gate.convertheight(), gate.convertwidth()];
                        return pricearray[gate.convertheight(), gate.convertwidth()];
                        break;

                    case "решетка 25х25":
                        for (int i = 0; i < 5; i++)
                            for (int j = 0; j < 13; j++)
                            {
                                Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[i + 41, j + 4];
                                pricearray[i, j] = Convert.ToInt32(cell.Value);
                            }
                        gate.price = pricearray[gate.convertheight(), gate.convertwidth()];
                        return pricearray[gate.convertheight(), gate.convertwidth()];
                        break;
                }
            }
            catch(System.NullReferenceException)
            {
                MessageBox.Show("Ошибка! Введите путь к документу Excel");
            }


            return 0;
        }

        public int GetWPrice()
        {
                switch (wicket.type)
                {
                    case "В полотне ворот":
                        return 10000;
                        break;
                    case "Отдельностоящая в собственной раме на проем":

                        if (wicket.width <= 1200 && wicket.height <= 2000)
                        {
                            if (wicket.execution == "рама без обшивки") return 10000;
                            if (wicket.execution == "проф.лист C-8 одна сторона") return 11600;
                            if (wicket.execution == "проф.лист C-8 две стороны") return 13200;
                            if (wicket.execution == "гладкий лист 2 мм") return 14800;
                        }

                        if (wicket.width <= 1200 && wicket.height > 2000 && wicket.height <= 2500)
                        {
                            if (wicket.execution == "рама без обшивки") return 11000;
                            if (wicket.execution == "проф.лист C-8 одна сторона") return 13000;
                            if (wicket.execution == "проф.лист C-8 две стороны") return 15000;
                            if (wicket.execution == "гладкий лист 2 мм") return 17000;
                        }

                    if (wicket.width <= 1200 && wicket.height > 2500 && wicket.height <= 3000)
                    {
                        if (wicket.execution == "рама без обшивки") return 12000;
                        if (wicket.execution == "проф.лист C-8 одна сторона") return 14400;
                        if (wicket.execution == "проф.лист C-8 две стороны") return 16800;
                        if (wicket.execution == "гладкий лист 2 мм") return 19200;
                    }

                    break;
                        
                }
            
            


            return 0;
        }

        public int GetAPrice()
        {
            switch (auto.acessories)
            {
                case "Nice": return 200;
                case " ": return 0;
            }
            return 0;
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
            try
            {
                excel.excelapp.Windows[1].Close(false);
            }
            catch (Exception)
            {
                
            }
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
        //Ворота____________________________________________________________________________________________
        private void Width_TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void Height_TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void Discount_Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void Discount_Text_Leave(object sender, EventArgs e)
        {
            if (Discount_Text_Gate.Text != "")
                gate.discount = Convert.ToInt32(Discount_Text_Gate.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }

            if (Discount_Text_Gate.Text != "")
                Discount_Text_Gate.Text = Discount_Text_Gate.Text + '%';

            if (gate.discount != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice() - Convert.ToInt32(gate.discount / 100 * GetPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
            }
        }

        private void Discount_Text_Enter(object sender, EventArgs e)
        {
            if (Discount_Text_Gate.Text != "")
            {
                int num = Discount_Text_Gate.Text.LastIndexOf("%");
                Discount_Text_Gate.Text = Discount_Text_Gate.Text.Remove(num);
                Discount_Text_Gate.SelectionStart = 0;
                Discount_Text_Gate.SelectionLength = Discount_Text_Gate.Text.Length;
                Discount_Text_Gate.Focus();

            }
        }
        //Код, отвечающий за изменение полей класса Gate
        private void OttdelkaBoxGate_TextChanged(object sender, EventArgs e)
        {
            gate.execution = OttdelkaBoxGate.Text;

            if (Discount_Text_Gate.Text != "")
                gate.discount = Convert.ToInt32(Discount_Text_Gate.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }

            if (Discount_Text_Gate.Text != "")
                Discount_Text_Gate.Text = Discount_Text_Gate.Text + '%';

            if (gate.discount != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice() - Convert.ToInt32(gate.discount / 100 * GetPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
            }
        }

        private void TypeBoxGate_TextChanged(object sender, EventArgs e)
        {
            gate.type = TypeBoxGate.Text;

            if (Discount_Text_Gate.Text != "")
                gate.discount = Convert.ToInt32(Discount_Text_Gate.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }

            if (Discount_Text_Gate.Text != "")
                Discount_Text_Gate.Text = Discount_Text_Gate.Text + '%';

            if (gate.discount != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice() - Convert.ToInt32(gate.discount / 100 * GetPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
            }
        }

        private void Width_TextBox_TextChanged(object sender, EventArgs e)
        {
            if (Discount_Text_Gate.Text != "")
                gate.discount = Convert.ToInt32(Discount_Text_Gate.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (Width_TextBox.Text != "")
                gate.width = Convert.ToInt32(Width_TextBox.Text);

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }

            if (Discount_Text_Gate.Text != "")
                Discount_Text_Gate.Text = Discount_Text_Gate.Text + '%';

            if (gate.discount != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice() - Convert.ToInt32(gate.discount / 100 * GetPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
            }
        }

        private void Height_TextBox_TextChanged(object sender, EventArgs e)
        {
            if (Discount_Text_Gate.Text != "")
                gate.discount = Convert.ToInt32(Discount_Text_Gate.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (Height_TextBox.Text != "")
                gate.height = Convert.ToInt32(Height_TextBox.Text);

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }

            if (Discount_Text_Gate.Text != "")
                Discount_Text_Gate.Text = Discount_Text_Gate.Text + '%';

            if (gate.discount != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice() - Convert.ToInt32(gate.discount / 100 * GetPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
            }
        }
        

        //Калитка____________________________________________________________________________________________________
        private void width_Wicket_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void height_Wicket_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void Discount_Wicket_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void WithoutWicket_Check_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (WithoutWicket_Check_Box.Checked == true)
            {
                TypeBoxWicket.Enabled = false;
                width_Wicket.Enabled = false;
                height_Wicket.Enabled = false;
                Discount_Wicket.Enabled = false;
                Total_With_Discount_Text_Wicket.Enabled = false;
                OttdelkaBox_Wicket.Enabled = false;
                FurnituraBox_Wicket.Enabled = false;
            }
            else
            {
                TypeBoxWicket.Enabled = true;
                width_Wicket.Enabled = true;
                height_Wicket.Enabled = true;
                Discount_Wicket.Enabled = true;
                Total_With_Discount_Text_Wicket.Enabled = true;
                OttdelkaBox_Wicket.Enabled = true;
                FurnituraBox_Wicket.Enabled = true;
            }
        }

        private void TypeBoxWicket_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedState = TypeBoxWicket.SelectedItem.ToString();
            wicket.type = selectedState;
            if (selectedState == "В полотне ворот")
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(10000) + " руб.";
                wicket.furnitura = "замок врезной и гарнитур нажимных ручек";
                width_Wicket.Enabled = false;
                height_Wicket.Enabled = false;
                OttdelkaBox_Wicket.Enabled = false;
                FurnituraBox_Wicket.Enabled = false;
            }
            else
            {
                width_Wicket.Enabled = true;
                height_Wicket.Enabled = true;
                OttdelkaBox_Wicket.Enabled = true;
                FurnituraBox_Wicket.Enabled = true;
            }
        }

        private void width_Wicket_TextChanged(object sender, EventArgs e)
        {
            wicket.width = Convert.ToInt32(width_Wicket.Text);

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
                wicket.price = GetWPrice();
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice() - Convert.ToInt32(wicket.discount / 100 * GetWPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
            }
        }

        private void height_Wicket_TextChanged(object sender, EventArgs e)
        {
            wicket.height = Convert.ToInt32(height_Wicket.Text);

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
                wicket.price = GetWPrice();
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice() - Convert.ToInt32(wicket.discount / 100 * GetWPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
            }
        }

        private void OttdelkaBox_Wicket_TextChanged(object sender, EventArgs e)
        {
            wicket.execution = OttdelkaBox_Wicket.Text;

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
                wicket.price = GetWPrice();
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice() - Convert.ToInt32(wicket.discount / 100 * GetWPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
            }
        }

        private void FurnituraBox_Wicket_TextChanged(object sender, EventArgs e)
        {
            wicket.furnitura = FurnituraBox_Wicket.Text;

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
                wicket.price = GetWPrice();
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice() - Convert.ToInt32(wicket.discount / 100 * GetWPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
            }
        }

        private void Discount_Wicket_Leave(object sender, EventArgs e)
        {
            if (Discount_Wicket.Text != "")
                wicket.discount = Convert.ToInt32(Discount_Wicket.Text);
            else
                wicket.discount = 0;

            if (Discount_Wicket.Text != "")
                Discount_Wicket.Text = Discount_Wicket.Text + '%';

            if (TypeBoxWicket.Text == "Отдельностоящая в собственной раме на проем") { 
                
                if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
                {
                    Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
                    wicket.price = GetWPrice();
                }
 
                if (wicket.discount != 0)
                {
                    Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice() - Convert.ToInt32(wicket.discount / 100 * GetWPrice())) + " руб.";
                }
                else
                {
                    Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
                }
            }
            else
            {
                wicket.price = 10000;
                if (wicket.discount != 0)
                {
                    Total_With_Discount_Text_Wicket.Text = Convert.ToString(10000 - Convert.ToInt32(wicket.discount / 100 * 10000)) + " руб.";
                }
                else
                {
                    Total_With_Discount_Text_Wicket.Text = Convert.ToString(10000) + " руб.";
                }
            }
        }

        private void Discount_Wicket_Enter(object sender, EventArgs e)
        {
            if (Discount_Wicket.Text != "")
            {
                int num = Discount_Wicket.Text.LastIndexOf("%");
                Discount_Wicket.Text = Discount_Wicket.Text.Remove(num);
                Discount_Wicket.SelectionStart = 0;
                Discount_Wicket.SelectionLength = Discount_Wicket.Text.Length;
                Discount_Wicket.Focus();

            }
        }
        //Автоматика_________________________________________________________________________________________________________
        private void Discount_Box_Auto_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void WithoutAuto_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (WithoutAuto_checkBox.Checked == true)
            {
                Producer_Box_Auto.Enabled = false;
                Acessories_Box_Auto.Enabled = false;
            }
            else
            {
                Producer_Box_Auto.Enabled = true;
                Acessories_Box_Auto.Enabled = true;
            }
        }

        private void Producer_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            auto.producer = Producer_Box_Auto.Text;

            switch (Producer_Box_Auto.Text)
            {
                case "Nice":
                    Acessories_Box_Auto.Items.Clear();
                    Acessories_Box_Auto.Items.AddRange(new string[] { "Nice", "" });
                    break;

                case "Came":
                    Acessories_Box_Auto.Items.Clear();
                    Acessories_Box_Auto.Items.AddRange(new string[] { "Came", "" });
                    break;

                case "FAAC":
                    Acessories_Box_Auto.Items.Clear();
                    Acessories_Box_Auto.Items.AddRange(new string[] { "Came", "" });
                    break;

                case "Doorhan":
                    Acessories_Box_Auto.Items.Clear();
                    Acessories_Box_Auto.Items.AddRange(new string[] { "Doorhan", "" });
                    break;

                case "An-Motors":
                    Acessories_Box_Auto.Items.Clear();
                    Acessories_Box_Auto.Items.AddRange(new string[] { "An-Motors", "" });
                    break;

                case "Comunello":
                    Acessories_Box_Auto.Items.Clear();
                    Acessories_Box_Auto.Items.AddRange(new string[] { "Comunello", "" });
                    break;

                case "Alutech":
                    Acessories_Box_Auto.Items.Clear();
                    Acessories_Box_Auto.Items.AddRange(new string[] { "Alutech", "" });
                    break;

            }
        }

        private void Discount_Box_Auto_Leave(object sender, EventArgs e)
        {   
            if (Discount_Box_Auto.Text != "")
                auto.discount = Convert.ToInt32(Discount_Box_Auto.Text);
            else
                auto.discount = 0;

            if (Discount_Box_Auto.Text != "")
                Discount_Box_Auto.Text = Discount_Box_Auto.Text + '%';
            
            if (Producer_Box_Auto.Text != "" && Acessories_Box_Auto.Text != "" )
            {
                Total_With_Discount_Box_Auto.Text = Convert.ToString(GetAPrice());
                auto.price = GetAPrice();
            }

            if (auto.discount != 0)
            {
                Total_With_Discount_Box_Auto.Text = Convert.ToString(GetAPrice() - Convert.ToInt32(auto.discount / 100 * GetAPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Box_Auto.Text = Convert.ToString(GetAPrice()) + " руб.";
            }
        }


        private void Discount_Box_Auto_Enter(object sender, EventArgs e)
        {
            if (Discount_Box_Auto.Text != "")
            {
                int num = Discount_Box_Auto.Text.LastIndexOf("%");
                Discount_Box_Auto.Text = Discount_Box_Auto.Text.Remove(num);
                Discount_Box_Auto.SelectionStart = 0;
                Discount_Box_Auto.SelectionLength = Discount_Box_Auto.Text.Length;
                Discount_Box_Auto.Focus();

            }
        }

        private void Acessories_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            auto.acessories = Acessories_Box_Auto.Text;

            if (auto.discount != 0)
            {
                Total_With_Discount_Box_Auto.Text = Convert.ToString(GetAPrice() - Convert.ToInt32(auto.discount / 100 * GetAPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Box_Auto.Text = Convert.ToString(GetAPrice()) + " руб.";
            }
        }

        public void путьДляExcelToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void Add_Gate_Button_Click(object sender, EventArgs e)
        {
            //Ворота
            if (Width_TextBox.Text != "")
            {
                string OutGate = "Ворота откатные " + " на проем " + Convert.ToString(gate.width) + " x " + Convert.ToString(gate.height);
                string OutEx = "Обшивка: " + gate.execution;
                Program.Send_GW_data(OutGate, OutEx, GetPrice(), gate.discount);
            }
            //

            //Автоматика
            if (WithoutAuto_checkBox.Checked == false)
            {
                Program.Senddata(auto.acessories, "комп." , 1, auto.price, auto.discount, excel.excelapp, excel.excelworksheet1);

                if (Zadvijka_checkBox.Checked == true)
                {   //доделать!!!!
                    Program.Senddata("Задвижка", "шт.", 1, 0, 0, excel.excelapp, excel.excelworksheet1);
                }

                if (Proushiny_checkBox.Checked == true)
                {   //доделать!!!!
                    Program.Senddata("Проушины", "шт.", 1, 0, 0, excel.excelapp, excel.excelworksheet1);
                }
            }
            else
            {
                if (Zadvijka_checkBox.Checked == true)
                {   //доделать!!!!
                    Program.Senddata("Задвижка", "шт.", 1, 0, 0, excel.excelapp, excel.excelworksheet1);
                }

                if (Proushiny_checkBox.Checked == true)
                {   //доделать!!!!
                    Program.Senddata("Проушины", "шт.", 1, 0, 0, excel.excelapp, excel.excelworksheet1);
                }
            } 
            //

            //Калитка
            string OutWicket;
            string WFurn;
            string WEx;
            if (WithoutWicket_Check_Box.Checked == false)
            {
                if (wicket.type == "В полотне ворот")
                {
                    OutWicket = "Калитка в полотне ворот";
                    WFurn = "Фурнитура: замок врезной и гарнитур нажимных ручек";
                    Program.Send_GW_data(OutWicket, WFurn, 10000, wicket.discount);
                }
                else
                {
                    OutWicket = "Калитка отдельностоящая в собственной раме на проем" + Convert.ToString(wicket.width) + " x " + Convert.ToString(wicket.height);
                    WEx = "Обшивка: " + wicket.execution;
                    Program.Send_GW_data(OutWicket, WEx, wicket.price, wicket.discount);

                    if (wicket.furnitura == "Гарнитур нажимных ручек, врезной замок и профильный цилиндр с тремя ключами")
                    {
                        WFurn = "Фурнитура: " + wicket.furnitura;
                        Program.Senddata(WFurn, "комп.", 1, 3000, 0, excel.excelapp, excel.excelworksheet1);
                    }
                    if (wicket.furnitura == "Электромеханическая запорная планка с врезным замком и ручкой-скобой")
                    {
                        WFurn = "Фурнитура: " + wicket.furnitura;
                        Program.Senddata(WFurn, "комп.", 1, 6000, 0, excel.excelapp, excel.excelworksheet1);
                    }

                }
            }
            //

            //Дополнительные позиции
            if (DopName1.Text != "" && DopCount1.Text != "" && DopPrice1.Text != "")
            {
                Program.Senddata(DopName1.Text, "", Convert.ToInt32(DopCount1.Text), Convert.ToInt32(DopPrice1.Text), 0, excel.excelapp, excel.excelworksheet1);
            }
            if (DopName2.Text != "" && DopCount2.Text != "" && DopPrice2.Text != "")
            {
                Program.Senddata(DopName2.Text, "", Convert.ToInt32(DopCount2.Text), Convert.ToInt32(DopPrice2.Text), 0, excel.excelapp, excel.excelworksheet1);
            }
            if (DopName3.Text != "" && DopCount3.Text != "" && DopPrice3.Text != "")
            {
                Program.Senddata(DopName3.Text, "", Convert.ToInt32(DopCount3.Text), Convert.ToInt32(DopPrice3.Text), 0, excel.excelapp, excel.excelworksheet1);
            }
            if (DopName4.Text != "" && DopCount4.Text != "" && DopPrice4.Text != "")
            {
                Program.Senddata(DopName4.Text, "", Convert.ToInt32(DopCount4.Text), Convert.ToInt32(DopPrice4.Text), 0, excel.excelapp, excel.excelworksheet1);
            }
            if (DopName5.Text != "" && DopCount5.Text != "" && DopPrice5.Text != "")
            {
                Program.Senddata(DopName5.Text, "", Convert.ToInt32(DopCount5.Text), Convert.ToInt32(DopPrice5.Text), 0, excel.excelapp, excel.excelworksheet1);
            }
            //

            excel.changevisible(true);

            MainPage pag = new MainPage();
            pag.Show();
            this.Hide();
        }

        private void Next_Button_Click(object sender, EventArgs e)
        {
        }
        
    }
    }
