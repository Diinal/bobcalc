﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Bobcalc
{
    public partial class OtkatPage : Form
    {
        Gate gate = new Gate();
        Wicket wicket = new Wicket();
        Auto auto = new Auto();
        DopPos doppos = new DopPos();
        Works works = new Works();
        bool checkstate;

        // static string path = "" /*"E:\\Bobmaster\\Bobmaster Calculator"*/;

        public OtkatPage()
        {
            InitializeComponent();
            this.DoubleBuffered = true;

            TypeBoxGate.Text = "Откатные";

            Microsoft.Win32.RegistryKey clientname = null;
            try
            {
                clientname = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                ClientName_textBox.Text = Convert.ToString(clientname.GetValue("ClientName"));
            }
            finally
            {
                if (clientname != null) clientname.Close();
            }

            Microsoft.Win32.RegistryKey clientemail = null;
            try
            {
                clientemail = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                ClientEmail_textBox.Text = Convert.ToString(clientemail.GetValue("ClientEmail"));
            }
            finally
            {
                if (clientemail != null) clientemail.Close();
            }

            Microsoft.Win32.RegistryKey clientphone = null;
            try
            {
                clientphone = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                ClientPhone_textBox.Text = Convert.ToString(clientphone.GetValue("ClientPhone"));
            }
            finally
            {
                if (clientphone != null) clientphone.Close();
            }
            excel.exw = excel.exw;
        }

        public int GetPrice()
        {
            try
            {
                int[,] pricearray = new int[5, 13];
                switch (gate.execution)
                {
                    case "рама без обшивки":
                        {   Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[gate.convertheight() + 4, gate.convertwidth() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);
                            break;
                        }
                    case "проф.лист C-8 одна сторона":
                        {
                            Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[gate.convertheight() + 13, gate.convertwidth() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);
                            break;
                        }
                    case "проф.лист C-8 две стороны":
                        {
                            Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[gate.convertheight() + 22, gate.convertwidth() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);
                            break;
                        }
                    case "гладкий лист 2 мм":
                        {
                            Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[gate.convertheight() + 31, gate.convertwidth() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);
                            break;
                        }
                    case "решетка 25х25":
                        {
                            Excel.Range cell = (Excel.Range)excel.excelworksheet2.Cells[gate.convertheight() + 41, gate.convertwidth() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);
                            break;
                        }
                }
            }
            catch (System.NullReferenceException)
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
                        if (wicket.execution == "без обшивки") return 10000;
                        if (wicket.execution == "C-8 одна сторона") return 11600;
                        if (wicket.execution == "C-8 две стороны") return 13200;
                        if (wicket.execution == "лист 2 мм") return 14800;
                    }

                    if (wicket.width <= 1200 && wicket.height > 2000 && wicket.height <= 2500)
                    {
                        if (wicket.execution == "без обшивки") return 11000;
                        if (wicket.execution == "C-8 одна сторона") return 13000;
                        if (wicket.execution == "C-8 две стороны") return 15000;
                        if (wicket.execution == "лист 2 мм") return 17000;
                    }

                    if (wicket.width <= 1200 && wicket.height > 2500 && wicket.height <= 3000)
                    {
                        if (wicket.execution == "без обшивки") return 12000;
                        if (wicket.execution == "C-8 одна сторона") return 14400;
                        if (wicket.execution == "C-8 две стороны") return 16800;
                        if (wicket.execution == "лист 2 мм") return 19200;
                    }
                    break;

            }

            return 0;
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
        
        //Код, отвечающий за изменение полей класса Gate
        private void OttdelkaBoxGate_TextChanged(object sender, EventArgs e)
        {
            gate.execution = OttdelkaBoxGate.Text;

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {         
                gate.discount = Convert.ToInt32(Global_Discount.Text); //Заполняет поле скидка в классе ворот
            }
            else
                gate.discount = 0;

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(0) + " руб.";
                gate.price = 0;
            }

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

            if (TypeBoxGate.Text == "Распашные")
            {
                RaspPage rp = new RaspPage();
                rp.Show();
                this.Hide();
            }

            else if (TypeBoxGate.Text == "Секционные")
            {
                SecPage sp = new SecPage();
                sp.Show();
                this.Hide();
            }
        }

        private void Width_TextBox_TextChanged(object sender, EventArgs e)
        {
            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
                gate.discount = Convert.ToInt32(Global_Discount.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (Width_TextBox.Text != "")
                gate.width = Convert.ToInt32(Width_TextBox.Text);

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(0) + " руб.";
                gate.price = 0;
            }
            
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
            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
                gate.discount = Convert.ToInt32(Global_Discount.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (Height_TextBox.Text != "")
                gate.height = Convert.ToInt32(Height_TextBox.Text);

            if (gate.execution != "" && gate.width != 0 && gate.height != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
                gate.price = GetPrice();
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(0) + " руб.";
                gate.price = 0;
            }

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
        

        private void TypeBoxWicket_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedState = TypeBoxWicket.SelectedItem.ToString();
            wicket.type = selectedState;
            if (selectedState == "В полотне ворот")
            {
                wicket.price_wick = 10000;
                if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
                {
                    Total_With_Discount_Text_Wicket.Text = Convert.ToString(10000 - (wicket.discount / 100 * 10000)) + " руб.";
                }
                else
                {
                    Total_With_Discount_Text_Wicket.Text = Convert.ToString(10000) + " руб.";
                }
                
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
            if (width_Wicket.Text != "")
                wicket.width = Convert.ToInt32(width_Wicket.Text);
            

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
                wicket.price_wick = GetWPrice();
            }
            else
            {
                wicket.price_wick = 0;
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total() - Convert.ToInt32(wicket.discount / 100 * wicket.Price_Total())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
            }
        }

        private void height_Wicket_TextChanged(object sender, EventArgs e)
        {
            if (height_Wicket.Text != "")
                wicket.height = Convert.ToInt32(height_Wicket.Text);

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
                wicket.price_wick = GetWPrice();
            }
            else
            {
                wicket.price_wick = 0;
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total() - Convert.ToInt32(wicket.discount / 100 * wicket.Price_Total())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
            }
        }

        private void OttdelkaBox_Wicket_TextChanged(object sender, EventArgs e)
        {
            wicket.execution = OttdelkaBox_Wicket.Text;

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
                wicket.price_wick = GetWPrice();
            }
            else
            {
                wicket.price_wick = 0;
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total() - Convert.ToInt32(wicket.discount / 100 * wicket.Price_Total())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
            }
        }

        private void FurnituraBox_Wicket_TextChanged(object sender, EventArgs e)
        {
            wicket.furnitura = FurnituraBox_Wicket.Text;

            if (wicket.furnitura == "Механическая")
            {
                wicket.price_fur = 3000;
            }
            else if (wicket.furnitura == "Электромеханическая")
            {
                wicket.price_fur = 6000;
            }
            else
            {
                wicket.price_fur = 0;
            }

            if (wicket.execution != "" && wicket.width != 0 && wicket.height != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
                wicket.price_wick = GetWPrice();
            }

            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total() - Convert.ToInt32(wicket.discount / 100 * wicket.Price_Total())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total()) + " руб.";
            }
        }
        
       
        
        //Автоматика_________________________________________________________________________________________________________
        private void Discount_Box_Auto_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }
        
        

        private void Producer_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            auto.producer = Producer_Box_Auto.Text;

            switch (Producer_Box_Auto.Text)
            {    
                case "Nice":
                    Drivers_Box_Auto.Items.Clear();
                    Drivers_Box_Auto.Items.AddRange(new string[] { "","RB600", "RB1000", "RUN1500", "RUN1800", "RUN2500", "RUN400HS", "RUN1200HS"});
                    Controls_comboBox.Items.Clear();
                    Controls_comboBox.Items.AddRange(new string[] { "", "FLO2RE", "FLO4RE","SM 2 RO 1","SM 4 RO 1" });
                    Photoelem_comboBox.Items.Clear();
                    Photoelem_comboBox.Items.AddRange(new string[] { "", "EPMB" });
                    Priemnik_comboBox.Items.Clear();
                    Priemnik_comboBox.Items.AddRange(new string[] { "", "SMXIS", "OXI" });
                    break;

                case "Came":
                    Drivers_Box_Auto.Items.Clear();
                    Drivers_Box_Auto.Items.AddRange(new string[] { "Came", "" });
                    break;

                case "FAAC":
                    Drivers_Box_Auto.Items.Clear();
                    Drivers_Box_Auto.Items.AddRange(new string[] { "", "740", "741", "746 ER", "844 ER", "844 R 3PH", "884 MC 3PH" });
                    Controls_comboBox.Items.Clear();
                    Controls_comboBox.Items.AddRange(new string[] { "", "XT2 868 SLH", "XT4 868 SLH"});
                    Photoelem_comboBox.Items.Clear();
                    Photoelem_comboBox.Items.AddRange(new string[] { "", "XP20W D", "XP20 D" });
                    Priemnik_comboBox.Items.Clear();
                    Priemnik_comboBox.Items.AddRange(new string[] { "", "RP 2 868" });
                    break;

                case "Doorhan":
                    Drivers_Box_Auto.Items.Clear();
                    Drivers_Box_Auto.Items.AddRange(new string[] { "Doorhan", "" });
                    break;

                case "An-Motors":
                    Drivers_Box_Auto.Items.Clear();
                    Drivers_Box_Auto.Items.AddRange(new string[] { "", "ASL 500", "ASL 1000", "ASL 2000" });
                    Controls_comboBox.Items.Clear();
                    Controls_comboBox.Items.AddRange(new string[] { "", "AT-4" });
                    Photoelem_comboBox.Items.Clear();
                    Photoelem_comboBox.Items.AddRange(new string[] { "", "P5103" });
                    break;

                case "Comunello":
                    Drivers_Box_Auto.Items.Clear();
                    Drivers_Box_Auto.Items.AddRange(new string[] { "", "FT 500", "FT 700", "FT 1000" });
                    Controls_comboBox.Items.Clear();
                    Controls_comboBox.Items.AddRange(new string[] { "", "KEEP-2", "KEEP-4" });
                    Photoelem_comboBox.Items.Clear();
                    Photoelem_comboBox.Items.AddRange(new string[] { "", "DTS" });
                    break;

                case "Alutech":
                    Drivers_Box_Auto.Items.Clear();
                    Drivers_Box_Auto.Items.AddRange(new string[] { "Alutech", "" });
                    break;

            }
        }

        private void ControlsCount_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            ControlsCount_Box_Auto.BackColor = Color.WhiteSmoke;

            if (ControlsCount_Box_Auto.Text != "")
            {
                switch (Controls_comboBox.Text)
                {
                    //Faac
                    case "":
                        auto.price_controls = 0;
                        break;
                    case "XT2 868 SLH":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 2033;
                        break;
                    case "XT4 868 SLH":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 2373;
                        break;
                    //Comunello
                    case "KEEP-2":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1000;
                        break;
                    case "KEEP-4":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1070;
                        break;
                    //AnMotors
                    case "AT-4":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 724;
                        break;
                    //Nice
                    case "FLO2RE":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1447;
                        break;
                    case "FLO4RE":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1591;
                        break;
                    case "SM 2 RO 1":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 965;
                        break;
                    case "SM 4 RO 1":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1076;
                        break;
                }

                if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
                {
                    TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
                }
                else
                {
                    TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
                }
            }

        }

        private void Rail_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            Rail_Box_Auto.BackColor = Color.WhiteSmoke;
            if (Rail_Box_Auto.Text != "")
                auto.price_rail = 720 * Convert.ToInt32(Rail_Box_Auto.Text);
            else
                auto.price_rail = 0;

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
            }
            else
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
            }
        }

        private void Drivers_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            switch (Drivers_Box_Auto.Text)
            {
                //FAAC
                case "":
                    auto.price_drive = 0;
                    break;
                case "740":
                    auto.price_drive = 15232;
                    break;
                case "741":
                    auto.price_drive = 20672;
                    break;
                case "746 ER":
                    auto.price_drive = 28729;
                    break;
                case "844 ER":
                    auto.price_drive = 31177;
                    break;
                case "844 R 3PH":
                    auto.price_drive = 61120;
                    break;
                case "884 MC 3PH":
                    auto.price_drive = 117310;
                    break;

                //Comunello
                case "FT 500":
                    auto.price_drive = 16125;
                    break;
                case "FT 700":
                    auto.price_drive = 18375;
                    break;
                case "FT 1000":
                    auto.price_drive = 22875;
                    break;

                //AnMotors
                case "ASL 500":
                    auto.price_drive = 12730;
                    break;
                case "ASL 1000":
                    auto.price_drive = 14841;
                    break;
                case "ASL 2000":
                    auto.price_drive = 19063;
                    break;
                //Nice
                case "RB600":
                    auto.price_drive = 18160;
                    break;
                case "RB1000":
                    auto.price_drive = 22160;
                    break;
                case "RUN1500":
                    auto.price_drive = 35920;
                    break;
                case "RUN1800":
                    auto.price_drive = 33200;
                    break;
                case "RUN2500":
                    auto.price_drive = 39240;
                    break;
                case "RUN400HS":
                    auto.price_drive = 38040;
                    break;
                case "RUN1200HS":
                    auto.price_drive = 40800;
                    break;


            }

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
            }
            else
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
            }
        }

        private void Controls_comboBox_TextChanged(object sender, EventArgs e)
        {
            if (ControlsCount_Box_Auto.Text != "")
            {
                switch (Controls_comboBox.Text)
                {
                    //Faac
                    case "":
                        auto.price_controls = 0;
                        break;
                    case "XT2 868 SLH":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 2033;
                        break;
                    case "XT4 868 SLH":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 2373;
                        break;
                    //Comunello
                    case "KEEP-2":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1000;
                        break;
                    case "KEEP-4":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1070;
                        break;
                    //AnMotors
                    case "AT-4":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 724;
                        break;
                    //Nice
                    case "FLO2RE":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1447;
                        break;
                    case "FLO4RE":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1591;
                        break;
                    case "SM 2 RO 1":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 965;
                        break;
                    case "SM 4 RO 1":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1076;
                        break;
                }

                if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
                {
                    TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
                }
                else
                {
                    TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
                }
            }
        }

        private void Priemnik_comboBox_TextChanged(object sender, EventArgs e)
        {
            switch (Priemnik_comboBox.Text)
            {
                case "":
                    auto.price_priemnik = 0;
                    break;
                case "SMXIS":
                    auto.price_priemnik = 1440;
                    break;
                case "OXI":
                    auto.price_priemnik = 2664;
                    break;
                case "RP 2 868":
                    auto.price_priemnik = 2856;
                    break;
            }


            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
            }
            else
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
            }
        }

        private void Photoelem_comboBox_TextChanged(object sender, EventArgs e)
        {
            switch (Photoelem_comboBox.Text)
            {
                case "":
                    auto.price_photoel = 0;
                    break;
                //FAAC
                case "XP20W D":
                    auto.price_photoel = 4420;
                    break;
                case "XP20 D":
                    auto.price_photoel = 3740;
                    break;
                //Comunello
                case "DTS":
                    auto.price_photoel = 2500;
                    break;
                //AnMotors
                case "P5103":
                    auto.price_photoel = 2500;
                    break;
                //Nice
                case "EPMB":
                    auto.price_photoel = 4032;
                    break;
            }

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
            }
            else
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
            }
        }

        private void Lamp_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Lamp_checkBox.Checked == true)
            {
                switch (Producer_Box_Auto.Text)
                {
                    case "":
                        auto.price_lamp = 0;
                        break;
                    case "Nice":
                        auto.price_lamp = 2232;
                        break;
                    case "FAAC":
                        auto.price_lamp = 2040;
                        break;
                    case "An - Motors":
                        auto.price_lamp = 1083;
                        break;
                    case "Comunello":
                        auto.price_lamp = 2150;
                        break;
                }
            }
            else
            {
                auto.price_lamp = 0;
            }

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
            }
            else
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
            }
        }
        
        public void путьДляExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string value = "";

            Microsoft.Win32.RegistryKey excelpath = null;
            try
            {
                excelpath = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                value = Convert.ToString(excelpath.GetValue("ExcelPath"));
            }
            catch { }
            
            if (Program.InputBox("Путь к Excel документу", "Введите путь к документу Excel:", ref value) == DialogResult.OK)
            {
                string path = value;
                //Microsoft.Win32.RegistryKey excelpath = null;
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
                MessageBox.Show("Пожалуйста, перезапустите приложение");
            }

        }

        private void путьДляPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string value = "";

            Microsoft.Win32.RegistryKey pdfpath = null;
            try
            {
                pdfpath = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                value = Convert.ToString(pdfpath.GetValue("PDFPath"));
            }
            catch { }

            if (Program.InputBox("Путь для PDF документа", "Введите путь для коммерческкого предложения:", ref value) == DialogResult.OK)
            {
                string path = value;
               // Microsoft.Win32.RegistryKey pdfpath = null;
                try
                {
                    pdfpath = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                    if (pdfpath != null)
                        pdfpath.SetValue("PDFPath", path);
                    pdfpath = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\VB and VBA Program Settings\\BobCalc\\PDF");
                    if (pdfpath != null)
                        pdfpath.SetValue("PDFPath", path);
                }
                finally
                {
                    if (pdfpath != null) pdfpath.Close();
                }

            }
        }

       /* public void ShowLoad()
        {
            load load = new load();
            load.Show();
            this.checkstate = false;
            Thread.Sleep(1000);
            while (!this.checkstate)
            {
            }
            Thread.Sleep(1000);
            load.Hide();
        }*/

        private void Add_Gate_Button_Click(object sender, EventArgs e)
        {
            bool keystart = true;
            int rubdiscount;

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "₽")
            {
                rubdiscount = Convert.ToInt32(Global_Discount.Text);
            }
            else
            {
                rubdiscount = 0;
            }

            //проверка инф о клиенте
            if (ClientName_textBox.Text == "")
            {
                keystart = false;
                ClientName_textBox.BackColor = Color.Salmon;
            }

            if (ClientEmail_textBox.Text == "")
            {
                keystart = false;
                ClientEmail_textBox.BackColor = Color.Salmon;
            }

            if (ClientPhone_textBox.Text == "")
            {
                keystart = false;
                ClientPhone_textBox.BackColor = Color.Salmon;
            }
            //

            //проверка автоматики
            if (Controls_comboBox.Text != "" && ControlsCount_Box_Auto.Text == "")
            {
                keystart = false;
                ControlsCount_Box_Auto.BackColor = Color.Salmon;
            }
            if (Drivers_Box_Auto.Text !="" && Rail_Box_Auto.Text == "" && Width_TextBox.Text != "")
            {
                keystart = false;
                Rail_Box_Auto.BackColor = Color.Salmon;
            }
            //

            //Проверка работы
            if (checkBox_AutoInst.Checked == true && Work_AutoInst.Text == "")
            {
                Work_AutoInst.BackColor = Color.Salmon;
                keystart = false;
            }
            if (checkBox_Delivery.Checked == true && Work_Delivery.Text == "")
            {
                Work_Delivery.BackColor = Color.Salmon;
                keystart = false;
            }
            if (checkBox_FoundBeton.Checked == true && Work_FoundBeton.Text == "")
            {
                Work_FoundBeton.BackColor = Color.Salmon;
                keystart = false;
            }
            if (checkBox_FoundVint.Checked == true && Work_FoundVint.Text == "")
            {
                Work_FoundVint.BackColor = Color.Salmon;
                keystart = false;
            }
            if (checkBox_GateInst.Checked == true && Work_GateInst.Text == "")
            {
                Work_GateInst.BackColor = Color.Salmon;
                keystart = false;
            }
            if (checkBox_PreWork.Checked == true && Work_PreWork.Text == "")
            {
                Work_PreWork.BackColor = Color.Salmon;
                keystart = false;
            }
            if (checkBox_WickInst.Checked == true && Work_WickInst.Text == "")
            {
                Work_WickInst.BackColor = Color.Salmon;
                keystart = false;
            }
            //
            if (keystart == true)
            {
             
                load ld = new Bobcalc.load();
                ld.Show();
                //Ворота
                if (Width_TextBox.Text != "")
                {
                    string OutGate = "Ворота откатные " + " на проем " + Convert.ToString(gate.width) + " x " + Convert.ToString(gate.height);
                    string OutEx = "Обшивка: " + gate.execution;
                    Program.Send_GW_data(OutGate, OutEx, GetPrice(), gate.discount, rubdiscount);
                }
                //

                //Автоматика
                if (Producer_Box_Auto.Text != "")
                {
                    

                    if (Drivers_Box_Auto.Text != "")
                    {
                        if (Rail_Box_Auto.Text != "")
                        {
                            Program.Senddata("Рейка зубчатая", "м.", Convert.ToInt32(Rail_Box_Auto.Text), 720, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                        }

                        switch (Drivers_Box_Auto.Text)
                        {
                            //FAAC
                            case "740":
                                Program.Senddata("Привод 740 Е, блок управления 740 D, 230В, Z16, V 12 м/мин, магнитные концевики, консистентная смазка, с монтажной пластиной ", "комп.", 1, 15232, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "741":
                                Program.Senddata("Привод 741 Е, блок управления 740 D, 230В, Z16, V 12 м/мин, магнитные концевики, консистентная смазка, с монтажной пластиной  ", "комп.", 1, 20672, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "746 ER":
                                Program.Senddata("Привод 746 ER, блок управления 780 D, 230В, Z16, V 9,5 м/мин, магнитные концевики, масляная ванна, без монтажной пластины  ", "комп.", 1, 27472, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Пластина монтажная с анкерами для приводов 746, 844 серий  ", "шт.", 1, 1257, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "844 ER":
                                Program.Senddata("Привод 844 ER, блок управления 780 D, 230В, Z16, V 9,5 м/мин, магнитные концевики, масляная ванна, без монтажной пластины ", "комп.", 1, 29920, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Пластина монтажная с анкерами для приводов 746, 844 серий  ", "шт.", 1, 1257, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "844 R 3PH":
                                Program.Senddata("Привод 844 R 3 фазы 380В, масляная ванна, без блока управления, без шестерни, без монтажной пластины", "комп.", 1, 39913, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Плата управления 844 Т для 1 мотора 380В  ", "шт.", 1, 16565, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Корпус 'L' для плат управления и принадлежностей, с замком под 3-гранный ключ, IP55, (ШxГxВ): 213x118x270мм  ", "шт.", 1, 1643, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Шестерня ведущая 20 зубьев модуль 4 для приводов 746, 844 серий, 12 м/мин", "шт.", 1, 1428, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Пластина монтажная с анкерами для приводов 746, 844 серий", "шт.", 1, 1571, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "884 MC 3PH":
                                Program.Senddata("Привод 884 MC 3 фазы, блок управления 884T 380В, масляная ванна, без шестерни, без монтажной пластины", "комп.", 1, 108800, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Пластина монтажная с анкерами для 884 MC", "шт.", 1, 5312, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                Program.Senddata("Шестерня ведущая 16 зубьев модуль 6 для привода 884, 10 м/мин", "шт.", 1, 3198, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;

                            //Comunello
                            case "FT 500":
                                Program.Senddata("Электропривод FT 500  со встроенным блоком управления; Радиоприемник;Два двухканальных пульта;Монтажный комплект", "комп.", 1, 16125, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "FT 700":
                                Program.Senddata("Электропривод FT 700  со встроенным блоком управления; Радиоприемник; Два двухканальных пульта;Монтажный комплект", "комп.", 1, 18375, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "FT 1000":
                                Program.Senddata("Электропривод  FT 1000 со встроенным блоком управления; Радиоприемник; Два двухканальных пульта;Монтажный комплект", "комп.", 1, 22875, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;

                            //AnMotors
                            case "ASL 500":
                                Program.Senddata("Электропривод ASL 500 со встроенным блоком управления; Встроенный радиоприемник;два четырехканальных пульта; монтажный комплект", "комп.", 1, 12730, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "ASL 1000":
                                Program.Senddata("Электропривод ASL 1000 со встроенным блоком управления; Встроенный радиоприемник;два четырехканальных пульта; монтажный комплект", "комп.", 1, 14841, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "ASL 2000":
                                Program.Senddata("Электропривод ASL 2000 со встроенным блоком управления; Встроенный радиоприемник;два четырехканальных пульта; монтажный комплект", "комп.", 1, 19063, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //Nice
                            case "RB600":
                                Program.Senddata("Привод для откатных ворот RB600", "шт.", 1, 18160, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RB1000":
                                Program.Senddata("Привод для откатных ворот RB1000", "шт.", 1, 22160, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RUN1500":
                                Program.Senddata("Привод для откатных ворот RUN1500", "шт.", 1, 35920, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RUN1800":
                                Program.Senddata("Привод для откатных ворот RUN1800", "шт.", 1, 33200, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RUN2500":
                                Program.Senddata("Привод для откатных ворот RUN2500", "шт.", 1, 39240, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RUN400HS":
                                Program.Senddata("Привод для откатных ворот RUN400HS", "шт.", 1, 38040, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RUN1200HS":
                                Program.Senddata("Привод для откатных ворот RUN1200HS", "шт.", 1, 40800, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;

                        }
                    }

                    if (Controls_comboBox.Text != "")
                    {
                        switch (Controls_comboBox.Text)
                        {
                            //Faac
                            case "XT2 868 SLH":
                                Program.Senddata("Брелок-передатчик XT2 868 SLH LR 868 МГц 2-канальный SLH код", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 2033, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "XT4 868 SLH":
                                Program.Senddata("Брелок-передатчик XT4 868 SLH LR 868 МГц 4-канальный SLH код", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 2373, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //Comunello
                            case "KEEP-2":
                                Program.Senddata("2-х канальный пульт дистанционного управления KEEP-2", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 1000, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "KEEP-4":
                                Program.Senddata("2-х канальный пульт дистанционного управления KEEP-4", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 1070, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //AnMotors
                            case "AT-4":
                                Program.Senddata("4-х канальный пульт дистанционного управления AT-4", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 724, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //Nice
                            case "FLO2RE":
                                Program.Senddata("2-х канальный пульт управления ERA FLOR FLO2RE", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 1447, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "FLO4RE":
                                Program.Senddata("2-х канальный пульт управления ERA FLOR FLO4RE", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 1591, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "SM 2 RO 1":
                                Program.Senddata("2-х канальный пульт управления SM 2 RO 1", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 965, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "SM 4 RO 1":
                                Program.Senddata("2-х канальный пульт управления SM 4 RO 1", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 1076, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                        }
                    }

                    if (Priemnik_comboBox.Text != "")
                    {
                        switch (Priemnik_comboBox.Text) {
                            case "SMXIS":
                                Program.Senddata("Приемник встраиваемый SMXIS", "шт.", 1, 1440, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "OXI":
                                Program.Senddata("Приемник встраиваемый OXI", "шт.", 1, 2664, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RP 2 868":
                                Program.Senddata("Приемник встраиваемый RP 2 868", "шт.", 1 , 2856, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                         }
                     }

                    if (Photoelem_comboBox.Text != "")
                    {
                        switch (Photoelem_comboBox.Text)
                        {
                            //FAAC
                            case "XP20W D":
                                Program.Senddata("Фотоэлементы XP20W D настенные, пара: приемник и передатчик c возможностью питания от батареи CR2", "шт.", 1, 4420, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "XP20 D":
                                Program.Senddata("Фотоэлементы XP20 D настенные, пара: приемник и передатчик", "шт.", 1, 3740, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //Comunello
                            case "DTS":
                                Program.Senddata("Фотоэлементы IR 30 проводные, компактные", "шт.", 1, 2500, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //AnMotors
                            case "P5103":
                                Program.Senddata("Фотоэлементы IR", "шт.", 1, 2500, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //Nice
                            case "EPMB":
                                Program.Senddata("Фотоэлементы Medium BlueBus EPMB", "шт.", 1, 4032, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                        }
                    }

                    if (Lamp_checkBox.Checked == true)
                    {
                        switch (Producer_Box_Auto.Text)
                        {
                            case "Nice":
                                Program.Senddata("Лампа сигнальная с антенной, 12В ELB", "шт.", 1, 2232, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "FAAC":
                                Program.Senddata("Лампа сигнальная FAACLIGHT, питание ~ 230В, 40Вт", "шт.", 1, 2040, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "An - Motors":
                                Program.Senddata("Сигнальная лампа F5002 230В с кронштейном крепления", "шт.", 1, 1083, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "Comunello":
                                Program.Senddata("Сигнальная лампа SWIFT светодиодная универсальная со встроенной антеной и кронштейном крепления", "шт.", 1, 2150, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                        }
                    }
                }

                if (Zadvijka_checkBox.Checked == true)
                {
                    Program.Senddata("Задвижка", "шт.", 1, 1000, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }

                if (Proushiny_checkBox.Checked == true)
                {
                    Program.Senddata("Проушины", "шт.", 1, 500, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }

                if (Handle_checkBox.Checked == true)
                {
                    Program.Senddata("Ручка", "шт.", 1, 2000, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                //

                //Калитка
                string OutWicket;
                string WFurn;
                string WEx;
                if (TypeBoxWicket.Text != "")
                {
                    if (wicket.type == "В полотне ворот")
                    {
                        OutWicket = "Калитка в полотне ворот";
                        WFurn = "Механическая фурнитура";
                        Program.Send_GW_data(OutWicket, WFurn, 10000, wicket.discount, rubdiscount);
                    }
                    else
                    {
                        OutWicket = "Калитка отдельностоящая в собственной раме на проем" + Convert.ToString(wicket.width) + " x " + Convert.ToString(wicket.height);
                        WEx = "Обшивка: " + wicket.execution;
                        Program.Send_GW_data(OutWicket, WEx, wicket.price_wick, wicket.discount, rubdiscount);

                        if (wicket.furnitura == "Механическая")
                        {
                            WFurn = "Механическая фурнитура";
                            Program.Senddata(WFurn, "комп.", 1, 3000, wicket.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                        }
                        if (wicket.furnitura == "Электромеханическая")
                        {
                            WFurn = "Электромеханическая фурнитура";
                            Program.Senddata(WFurn, "комп.", 1, 6000, wicket.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                        }

                    }
                }
                //

                //Дополнительные позиции
                if (DopName1.Text != "" && DopCount1.Text != "" && DopPrice1.Text != "")
                {
                    Program.Senddata(DopName1.Text, "", Convert.ToInt32(DopCount1.Text), Convert.ToInt32(DopPrice1.Text), doppos.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (DopName2.Text != "" && DopCount2.Text != "" && DopPrice2.Text != "")
                {
                    Program.Senddata(DopName2.Text, "", Convert.ToInt32(DopCount2.Text), Convert.ToInt32(DopPrice2.Text), doppos.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (DopName3.Text != "" && DopCount3.Text != "" && DopPrice3.Text != "")
                {
                    Program.Senddata(DopName3.Text, "", Convert.ToInt32(DopCount3.Text), Convert.ToInt32(DopPrice3.Text), doppos.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (DopName4.Text != "" && DopCount4.Text != "" && DopPrice4.Text != "")
                {
                    Program.Senddata(DopName4.Text, "", Convert.ToInt32(DopCount4.Text), Convert.ToInt32(DopPrice4.Text), doppos.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (DopName5.Text != "" && DopCount5.Text != "" && DopPrice5.Text != "")
                {
                    Program.Senddata(DopName5.Text, "", Convert.ToInt32(DopCount5.Text), Convert.ToInt32(DopPrice5.Text), doppos.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                //

                //Работа
                if (checkBox_AutoInst.Checked == true && Work_AutoInst.Text != "")
                {
                    Program.Senddata("Установка автоматики", "", 1, Convert.ToInt32(Work_AutoInst.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (checkBox_Delivery.Checked == true && Work_Delivery.Text != "")
                {
                    Program.Senddata("Доставка", "", 1, Convert.ToInt32(Work_Delivery.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (checkBox_FoundBeton.Checked == true && Work_FoundBeton.Text != "")
                {
                    Program.Senddata("Установка бетонного основания", "", 1, Convert.ToInt32(Work_FoundBeton.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (checkBox_FoundVint.Checked == true && Work_FoundVint.Text != "")
                {
                    Program.Senddata("Установка винтового основания", "", 1, Convert.ToInt32(Work_FoundVint.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (checkBox_GateInst.Checked == true && Work_GateInst.Text != "")
                {
                    Program.Senddata("Установка ворот", "", 1, Convert.ToInt32(Work_GateInst.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (checkBox_PreWork.Checked == true && Work_PreWork.Text != "")
                {
                    Program.Senddata("Подготовительные работы", "", 1, Convert.ToInt32(Work_PreWork.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                if (checkBox_WickInst.Checked == true && Work_WickInst.Text != "")
                {
                    Program.Senddata("Установка калитки", "", 1, Convert.ToInt32(Work_WickInst.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                //

                Excel.Range cell = (Excel.Range)excel.excelworksheet1.Cells[7, 3];
                cell.Value = ClientName_textBox.Text;

                cell = (Excel.Range)excel.excelworksheet1.Cells[8, 3];
                cell.Value = ClientPhone_textBox.Text;

                cell = (Excel.Range)excel.excelworksheet1.Cells[9, 3];
                cell.Value = ClientEmail_textBox.Text;
                
                OtkatPage pag = new OtkatPage();
                pag.Show();
                ld.Close();
                this.Hide();
            }
        }

        private void GoToKP_Button_Click(object sender, EventArgs e)
        {
            string value = "";
            if (Program.InputBox("Номер КП", "Введите номер коммерческого предложения:", ref value) == DialogResult.OK)
                excel.cellkp = excel.excelworksheet1.get_Range("B6");
            excel.cellkp.Value = "    Коммерческое предложение № " + value + " от " + DateTime.Today.ToShortDateString();
            excel.changevisible(true);
        }

        private void MainPage_FormClosed(object sender, FormClosedEventArgs e)
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

        //Массовая скидка
        private void Global_Discount_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
        }
        
        private void Discount_Type_comboBox_TextChanged(object sender, EventArgs e)
        {
            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                gate.discount = Convert.ToDouble(Global_Discount.Text);
                wicket.discount = Convert.ToDouble(Global_Discount.Text);
                doppos.discount = Convert.ToDouble(Global_Discount.Text);
                works.discount = Convert.ToDouble(Global_Discount.Text);
                auto.discount = Convert.ToDouble(Global_Discount.Text);
            }
            else
            {
                gate.discount = 0;
                wicket.discount = 0;
                doppos.discount = 0;
                works.discount = 0;
                auto.discount = 0;
            }
            if (gate.discount != 0)
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice() - Convert.ToInt32(gate.discount / 100 * GetPrice())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";
            }
            if (wicket.discount != 0)
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(wicket.Price_Total() - Convert.ToInt32(wicket.discount / 100 * wicket.Price_Total())) + " руб.";
            }
            else
            {
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";
            }
            if (auto.discount != 0)
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total() - Convert.ToInt32(auto.discount / 100 * auto.Price_Total())) + " руб.";
            }
            else
            {
                TotalAuto_textBox.Text = Convert.ToString(auto.Price_Total()) + " руб.";
            }


            if (Discount_Type_comboBox.Text == "₽")
            {
                gate.discount = 0;
                Total_With_Discount_Text_Gate.Text = Convert.ToString(GetPrice()) + " руб.";

                wicket.discount = 0;
                Total_With_Discount_Text_Wicket.Text = Convert.ToString(GetWPrice()) + " руб.";

                auto.discount = 0;

                doppos.discount = 0;

                works.discount = 0;
            }
        }
        //возврат дефолтных цветов при вводе текста
        private void Work_AutoInst_TextChanged(object sender, EventArgs e)
        {
            Work_AutoInst.BackColor = Color.WhiteSmoke;
        }

        private void Work_PreWork_TextChanged(object sender, EventArgs e)
        {
            Work_PreWork.BackColor = Color.WhiteSmoke;
        }

        private void Work_FoundBeton_TextChanged(object sender, EventArgs e)
        {
            Work_FoundBeton.BackColor = Color.WhiteSmoke;
        }

        private void Work_FoundVint_TextChanged(object sender, EventArgs e)
        {
            Work_FoundVint.BackColor = Color.WhiteSmoke;
        }

        private void Work_Delivery_TextChanged(object sender, EventArgs e)
        {
            Work_Delivery.BackColor = Color.WhiteSmoke;
        }

        private void Work_GateInst_TextChanged(object sender, EventArgs e)
        {
            Work_GateInst.BackColor = Color.WhiteSmoke;
        }

        private void Work_WickInst_TextChanged(object sender, EventArgs e)
        {
            Work_WickInst.BackColor = Color.WhiteSmoke;
        }

        private void ClientName_textBox_TextChanged(object sender, EventArgs e)
        {
            ClientName_textBox.BackColor = Color.WhiteSmoke;
            //Запись имени клиента в реестр для сохранения при обновлении формы в текущем заказе
            Microsoft.Win32.RegistryKey clientname = null;
            try
            {
                clientname = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                if (clientname != null)
                    clientname.SetValue("ClientName", ClientName_textBox.Text);
            }
            finally
            {
                if (clientname != null) clientname.Close();
            }
        }

        private void ClientEmail_textBox_TextChanged(object sender, EventArgs e)
        {
            ClientEmail_textBox.BackColor = Color.WhiteSmoke;
            //Запись почты клиента в реестр для сохранения при обновлении формы в текущем заказе
            Microsoft.Win32.RegistryKey clientemail = null;
            try
            {
                clientemail = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                if (clientemail != null)
                    clientemail.SetValue("ClientEmail", ClientEmail_textBox.Text);
            }
            finally
            {
                if (clientemail != null) clientemail.Close();
            }
        }

        private void ClientPhone_textBox_TextChanged(object sender, EventArgs e)
        {
            ClientPhone_textBox.BackColor = Color.WhiteSmoke;
            //Запись номера клиента в реестр для сохранения при обновлении формы в текущем заказе
            Microsoft.Win32.RegistryKey clientphone = null;
            try
            {
                clientphone = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                if (clientphone != null)
                    clientphone.SetValue("ClientPhone", ClientPhone_textBox.Text);
            }
            finally
            {
                if (clientphone != null) clientphone.Close();
            }
        }

        private void Start_button_Click(object sender, EventArgs e)
        {
            Microsoft.Win32.RegistryKey clientname = null;
            try
            {
                clientname = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                if (clientname != null)
                    clientname.SetValue("ClientName", "");
                    clientname.SetValue("ClientEmail", "");
                    clientname.SetValue("ClientPhone", "");
            }
            finally
            {
                if (clientname != null) clientname.Close();
            }

            OtkatPage mp = new OtkatPage();
            mp.Show();
            this.Hide();
        }

        



        //
    }
}
