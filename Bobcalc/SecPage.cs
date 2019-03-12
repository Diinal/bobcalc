using System;
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
    public partial class SecPage : Form
    {
        Gate gate = new Gate();
        Wicket wicket = new Wicket();
        Auto auto = new Auto();
        DopPos doppos = new DopPos();
        Works works = new Works();
        bool checkstate;

        public SecPage()
        {
            InitializeComponent();
            TypeBoxGate.Text = "Секционные";

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
        }

        public int GetPrice()
        {
            int[,] pricearray;


            if (TypeSecComboBox.Text == "Classic")
            {
                //pricearray = new int[12, 35];

                excel.exw = excel.excelworksheet4;
                
            }

            if (TypeSecComboBox.Text == "Trend")
            {
                //pricearray = new int[11, 35];

                excel.exw = excel.excelworksheet5;
            }
            
            try
                {                
                    switch (gate.execution)
                    {
                        case "Стандартно окр.":
                        {
                            Excel.Range cell = (Excel.Range)excel.exw.Cells[gate.convertheightsec(TypeSecComboBox.Text) + 4, gate.convertwidthsec() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);

                            break;
                        }
                        case "Пленка":
                        {
                            Excel.Range cell = (Excel.Range)excel.exw.Cells[gate.convertheightsec(TypeSecComboBox.Text) + 20, gate.convertwidthsec() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);
                            break;
                        }
                        case "Филенка стнд.":
                        {
                            Excel.Range cell = (Excel.Range)excel.exw.Cells[gate.convertheightsecfil() + 36, gate.convertwidthsecfil() + 4];

                            gate.price = Convert.ToInt32(cell.Value);
                            return Convert.ToInt32(cell.Value);
                            break;
                        }
                        case "Филенка пленка":
                        {
                            Excel.Range cell = (Excel.Range)excel.exw.Cells[gate.convertheightsecfil() + 54, gate.convertwidthsecfil() + 4];

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
        //Калитка
        private void Wicket_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (this.Wicket_checkBox.Checked)
            {
                if (this.wicket.discount != 0.0)
                {
                    if (this.TypeSecComboBox.Text == "Classic")
                    {
                        this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(35489 - Convert.ToInt32(this.wicket.discount / 100.0 * 35489.0)) + " руб.";
                    }
                    if (this.TypeSecComboBox.Text == "Trend")
                    {
                        this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(28702 - Convert.ToInt32(this.wicket.discount / 100.0 * 28702.0)) + " руб.";
                        return;
                    }
                }
                else if (this.wicket.discount == 0.0 && this.Wicket_checkBox.Checked)
                {
                    if (this.TypeSecComboBox.Text == "Classic")
                    {
                        this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(35489) + " руб.";
                    }
                    if (this.TypeSecComboBox.Text == "Trend")
                    {
                        this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(28702) + " руб.";
                        return;
                    }
                }
            }
            else
            {
                this.Total_With_Discount_Text_Wicket.Text = "";
            }
        }

        private void OnlyNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 59) && e.KeyChar != 8)
                e.Handled = true;
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

            SecPage sp = new SecPage();
            sp.Show();
            this.Hide();
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

        private void RaspPage_FormClosed(object sender, FormClosedEventArgs e)
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

        //Ворота_______________________________________________________________________________________
        private void TypeBoxGate_TextChanged(object sender, EventArgs e)
        {
            gate.type = TypeBoxGate.Text;

            if (TypeBoxGate.Text == "Откатные")
            {
                OtkatPage op = new OtkatPage();
                op.Show();
                this.Hide();
            }
            else if (TypeBoxGate.Text == "Распашные")
            {
                RaspPage rp = new RaspPage();
                rp.Show();
                this.Hide();
            }

        }

        private void TypeSecComboBox_TextChanged(object sender, EventArgs e)
        {
            gate.typesec = TypeSecComboBox.Text;

            if (TypeSecComboBox.Text == "Classic") excel.exw = excel.excelworksheet4;
            if (TypeSecComboBox.Text == "Trend") excel.exw = excel.excelworksheet5;

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                gate.discount = Convert.ToInt32(Global_Discount.Text); //Заполняет поле скидка в классе ворот
            }
            else
                gate.discount = 0;

            if (gate.typesec != "" && gate.execution != "" && gate.width != 0 && gate.height != 0)
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

            if (this.wicket.discount != 0.0 && this.Wicket_checkBox.Checked)
            {
                if (this.TypeSecComboBox.Text == "Classic")
                {
                    this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(35489 - Convert.ToInt32(0)) + " руб.";
                }
                if (this.TypeSecComboBox.Text == "Trend")
                {
                    this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(28702 - Convert.ToInt32(0)) + " руб.";
                    return;
                }
            }
            else if (this.wicket.discount == 0.0 && this.Wicket_checkBox.Checked)
            {
                if (this.TypeSecComboBox.Text == "Classic")
                {
                    this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(35489) + " руб.";
                }
                if (this.TypeSecComboBox.Text == "Trend")
                {
                    this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(28702) + " руб.";
                    return;
                }
            }
            else
            {
                this.Total_With_Discount_Text_Wicket.Text = "";
            }
        }

        private void OttdelkaBoxGate_TextChanged(object sender, EventArgs e)
        {
            gate.execution = OttdelkaBoxGate.Text;

            //TypeSecComboBox.Text;

            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
            {
                gate.discount = Convert.ToInt32(Global_Discount.Text); //Заполняет поле скидка в классе ворот
            }
            else
                gate.discount = 0;

            if (gate.typesec != "" && gate.execution != "" && gate.width != 0 && gate.height != 0)
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

        private void Width_TextBox_TextChanged(object sender, EventArgs e)
        {
            if (Global_Discount.Text != "" && Discount_Type_comboBox.Text == "%")
                gate.discount = Convert.ToInt32(Global_Discount.Text); //Заполняет поле скидка в классе ворот
            else
                gate.discount = 0;

            if (Width_TextBox.Text != "")
                gate.width = Convert.ToInt32(Width_TextBox.Text);

            if (gate.typesec != "" && gate.typesec != "" && gate.execution != "" && gate.width != 0 && gate.height != 0)
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

            if (gate.typesec != "" && gate.execution != "" && gate.width != 0 && gate.height != 0)
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

        

        //Автоматика_________________________________________________________________________________________________________

        private void Producer_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            auto.producer = Producer_Box_Auto.Text;
            if (this.Producer_Box_Auto.Text == "")
            {
                this.TotalAuto_textBox.Text = "";
            }
            string text = this.Producer_Box_Auto.Text;
            if (text == "Nice")
            {
                this.Drivers_Box_Auto.Items.Clear();
                this.Drivers_Box_Auto.Items.AddRange(new string[]{"", "Spin 21","Spin 22"});

                this.Controls_comboBox.Items.Clear();
                this.Controls_comboBox.Items.AddRange(new string[] {"","FLO2RE"});

                this.Photoelem_comboBox.Items.Clear();
                this.Photoelem_comboBox.Items.AddRange(new string[]{"","EPMB"});

                this.Priemnik_comboBox.Items.Clear();
                this.Priemnik_comboBox.Items.AddRange(new string[] { "", "SMXIS"});

                return;
            }
            
            if (text == "FAAC")
            {
                this.Drivers_Box_Auto.Items.Clear();
                this.Drivers_Box_Auto.Items.AddRange(new string[]{"", "D 100"});

                this.Controls_comboBox.Items.Clear();
                this.Controls_comboBox.Items.AddRange(new string[] {"","XT2 868 SLH","XT4 868 SLH"});

                this.Photoelem_comboBox.Items.Clear();
                this.Photoelem_comboBox.Items.AddRange(new string[]{ "", "XP20W D","XP20 D"});

                this.Priemnik_comboBox.Items.Clear();
                this.Priemnik_comboBox.Items.AddRange(new string[]{"","RP 2 868"});

                return;
            }
            if (text == "An-Motors")
            {
                this.Drivers_Box_Auto.Items.Clear();
                this.Drivers_Box_Auto.Items.AddRange(new string[]{"", "ASI 50 KIT"});

                this.Controls_comboBox.Items.Clear();
                this.Controls_comboBox.Items.AddRange(new string[]{ "","AT-4"});

                this.Photoelem_comboBox.Items.Clear();
                this.Photoelem_comboBox.Items.AddRange(new string[]{"","P5103"});

                return;
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
                    /*case "FLO4RE":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1591;
                        break;
                    case "SM 2 RO 1":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 965;
                        break;
                    case "SM 4 RO 1":
                        auto.price_controls = Convert.ToInt32(ControlsCount_Box_Auto.Text) * 1076;
                        break;*/
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


        private void Drivers_Box_Auto_TextChanged(object sender, EventArgs e)
        {
            switch (Drivers_Box_Auto.Text)
            {
                //FAAC
                case "":
                    auto.price_drive = 0;
                    break;
                case "D 100":
                    auto.price_drive = 18950;
                    break;

                //Nice
                case "Spin 21":
                    auto.price_drive = 13040;
                    break;
                case "Spin 22":
                    auto.price_drive = 14600;
                    break;

                //An-Motors
                case "ASI 50 KIT":
                    auto.price_drive = 30220;
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
                /*case "OXI":
                    auto.price_priemnik = 2664;
                    break;*/
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
                value = Convert.ToString(pdfpath.GetValue("ExcelPath"));
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

        /*public void ShowLoad()
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
            
            //
            if (keystart == true)
            {
                load ld = new Bobcalc.load();
                ld.Show();
                //Ворота
                if (Width_TextBox.Text != "")
                {   
                    string OutGate = "Ворота секционные " + TypeSecComboBox.Text + " на проем " + Convert.ToString(gate.width) + " x " + Convert.ToString(gate.height);
                    string OutEx = "Обшивка: " + gate.execution;
                    Program.Send_GW_data(OutGate, OutEx, GetPrice(), gate.discount, rubdiscount);
                }
                //

                //Автоматика
                if (Producer_Box_Auto.Text != "")
                {


                    if (Drivers_Box_Auto.Text != "")
                    {


                        switch (Drivers_Box_Auto.Text)
                        {
                            //FAAC
                            case "D100":
                                Program.Senddata("Привод D100, приемник, пульт, рельс", "комп.", 1, 18950, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;


                            //Nice
                            case "Spin 21":
                                Program.Senddata("Привод Spin 21, приемник, пульт, рельс", "комп.", 1, 13040, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "Spin 22":
                                Program.Senddata("Привод Spin 22, приемник, пульт, рельс", "комп.", 1, 14600, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;

                            //An-motors
                            case "ASI 50 KIT":
                                Program.Senddata("Привод ASI 50 KIT, кнопочный пост, блок управления, цепь ручного подъема", "комп.", 1, 30220, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
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
                            //AnMotors
                            case "AT-4":
                                Program.Senddata("4-х канальный пульт дистанционного управления AT-4", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 724, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            //Nice
                            case "FLO2RE":
                                Program.Senddata("2-х канальный пульт управления ERA FLOR FLO2RE", "шт.", Convert.ToInt32(ControlsCount_Box_Auto.Text), 1447, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                        }
                    }

                    if (Priemnik_comboBox.Text != "")
                    {
                        switch (Priemnik_comboBox.Text)
                        {
                            case "SMXIS":
                                Program.Senddata("Приемник встраиваемый SMXIS", "шт.", 1, 1440, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "OXI":
                                Program.Senddata("Приемник встраиваемый OXI", "шт.", 1, 2664, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                                break;
                            case "RP 2 868":
                                Program.Senddata("Приемник встраиваемый RP 2 868", "шт.", 1, 2856, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
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

                if (Reductor_checkBox.Checked == true)
                {
                    Program.Senddata("Редуктор с цепью", "шт.", 1, 7040, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }

                if (BlokRP_checkBox.Checked == true)
                {
                    Program.Senddata("Блок ручного управления", "шт.", 1, 1500, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }

                if (Zamok_checkBox.Checked == true)
                {
                    Program.Senddata("Замок ригельный", "шт.", 1, 6067, auto.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                }
                //

                //Калитка               
                 if (this.Wicket_checkBox.Checked)
                 {
                     if (this.TypeSecComboBox.Text == "Classic")
                     {
                         Program.Senddata("Калитка встроенная", "шт.", 1, 35489, this.wicket.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                     }
                     if (this.TypeSecComboBox.Text == "Trend")
                     {
                         Program.Senddata("Калитка встроенная", "шт.", 1, 28702, this.wicket.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
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

                 if (checkBox_GateInst.Checked == true && Work_GateInst.Text != "")
                 {
                     Program.Senddata("Установка ворот", "", 1, Convert.ToInt32(Work_GateInst.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                 }
                 if (checkBox_PreWork.Checked == true && Work_PreWork.Text != "")
                 {
                     Program.Senddata("Подготовительные работы", "", 1, Convert.ToInt32(Work_PreWork.Text), works.discount, excel.excelapp, excel.excelworksheet1, rubdiscount);
                 }
                 //

                 Excel.Range cell = (Excel.Range)excel.excelworksheet1.Cells[7, 3];
                 cell.Value = ClientName_textBox.Text;

                 cell = (Excel.Range)excel.excelworksheet1.Cells[8, 3];
                 cell.Value = ClientPhone_textBox.Text;

                 cell = (Excel.Range)excel.excelworksheet1.Cells[9, 3];
                 cell.Value = ClientEmail_textBox.Text;
                
                 SecPage pag = new SecPage();
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
                if (this.Wicket_checkBox.Checked)
                {
                    if (this.TypeSecComboBox.Text == "Classic")
                    {
                        this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(35489 - Convert.ToInt32(wicket.discount / 100 * 35489)) + " руб.";
                    }
                    if (this.TypeSecComboBox.Text == "Trend")
                    {
                        this.Total_With_Discount_Text_Wicket.Text = Convert.ToString(28702 - Convert.ToInt32(wicket.discount / 100 * 28702)) + " руб.";
                        return;
                    }
                }
                else
                {
                    if (this.TypeSecComboBox.Text == "Classic")
                    {
                        Total_With_Discount_Text_Wicket.Text = Convert.ToString(35489) + " руб.";
                    }
                    if (this.TypeSecComboBox.Text == "Trend")
                    {
                        Total_With_Discount_Text_Wicket.Text = Convert.ToString(28702) + " руб.";
                        return;
                    }
                }
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
                if (this.Wicket_checkBox.Checked)
                {
                    if (this.TypeSecComboBox.Text == "Classic")
                    {
                        Total_With_Discount_Text_Wicket.Text = Convert.ToString(35489) + " руб.";
                    }
                    if (this.TypeSecComboBox.Text == "Trend")
                    {
                        Total_With_Discount_Text_Wicket.Text = Convert.ToString(28702) + " руб.";
                        return;
                    }
                }
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
        

        private void Work_Delivery_TextChanged(object sender, EventArgs e)
        {
            Work_Delivery.BackColor = Color.WhiteSmoke;
        }

        private void Work_GateInst_TextChanged(object sender, EventArgs e)
        {
            Work_GateInst.BackColor = Color.WhiteSmoke;
        }
        

    }

}
