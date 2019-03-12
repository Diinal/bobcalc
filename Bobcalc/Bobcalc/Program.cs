using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace Bobcalc
{   
    class Gate
    {   //Поля
        public string type;
        public string execution;
        public int width;
        public int height;
        public double discount;
        public int price;

        //Методы
        public Gate() { }
        ~Gate() { }

        public int convertwidth()
        {
                if (width == 3000) return 0;
                else if (3001 <=width && width<=3500) return 1;
                else if (3501 <=width && width <= 4000) return 2;
                else if (4001 <= width && width <= 4500) return 3;
                else if (4501 <= width && width <= 5000) return 4;
                else if (5001 <= width && width <= 5500) return 5;
                else if (5501 <= width && width <= 6000) return 6;
                else if (6001 <= width && width <= 6500) return 7;
                else if (6501 <= width && width <= 7000) return 8;
                else if (7001 <= width && width <= 7500) return 9;
                else if (7501 <= width && width <= 8000) return 10;
                else if (8001 <= width && width <= 8500) return 11;
                else if (8501 <= width && width <= 9000) return 12;
                else return 0;
           
        }

        public int convertheight()
        {
            if (height ==2000) return 0;
            else if (2001 <= height && height <= 2250) return 1;
            else if (2251 <= height && height <= 2500) return 2;
            else if (2501 <= height && height <= 2750) return 3;
            else if (2751 <= height && height <= 3000) return 4;
            else return 0;
        }
   
        
       

    }

    class Wicket
    {
        //Поля
        public string type;
        public string execution;
        public string furnitura;
        public int width;
        public int height;
        public double discount;
        public int price;

        //Методы
        public Wicket() { }
        ~Wicket() { }
    }

    class Auto
    {
        //Поля
        public string producer;
        public string acessories;
        public double discount;
        public int price;

        //Методы
        public Auto() { }
        ~Auto() { }
    }

    public class NumByWords
    {
        public static string RurPhrase(decimal money)
        {
            return CurPhrase(money, "рубль", "рубля", "рублей", "копейка", "копейки", "копеек");
        }

        public static string UsdPhrase(decimal money)
        {
            return CurPhrase(money, "доллар США", "доллара США", "долларов США", "цент", "цента", "центов");
        }

        public static string NumPhrase(ulong Value, bool IsMale)
        {
            if (Value == 0UL) return "Ноль";
            string[] Dek1 = { "", " од", " дв", " три", " четыре", " пять", " шесть", " семь", " восемь", " девять", " десять", " одиннадцать", " двенадцать", " тринадцать", " четырнадцать", " пятнадцать", " шестнадцать", " семнадцать", " восемнадцать", " девятнадцать" };
            string[] Dek2 = { "", "", " двадцать", " тридцать", " сорок", " пятьдесят", " шестьдесят", " семьдесят", " восемьдесят", " девяносто" };
            string[] Dek3 = { "", " сто", " двести", " триста", " четыреста", " пятьсот", " шестьсот", " семьсот", " восемьсот", " девятьсот" };
            string[] Th = { "", "", " тысяч", " миллион", " миллиард", " триллион", " квадрилион", " квинтилион" };
            string str = "";
            for (byte th = 1; Value > 0; th++)
            {
                ushort gr = (ushort)(Value % 1000);
                Value = (Value - gr) / 1000;
                if (gr > 0)
                {
                    byte d3 = (byte)((gr - gr % 100) / 100);
                    byte d1 = (byte)(gr % 10);
                    byte d2 = (byte)((gr - d3 * 100 - d1) / 10);
                    if (d2 == 1) d1 += (byte)10;
                    bool ismale = (th > 2) || ((th == 1) && IsMale);
                    str = Dek3[d3] + Dek2[d2] + Dek1[d1] + EndDek1(d1, ismale) + Th[th] + EndTh(th, d1) + str;
                };
            };
            str = str.Substring(1, 1).ToUpper() + str.Substring(2);
            return str;
        }

        #region Private members
        private static string CurPhrase(decimal money,
            string word1, string word234, string wordmore,
            string sword1, string sword234, string swordmore)
        {
            money = decimal.Round(money, 2);
            decimal decintpart = decimal.Truncate(money);
            ulong intpart = decimal.ToUInt64(decintpart);
            string str = NumPhrase(intpart, true) + " ";
            byte endpart = (byte)(intpart % 100UL);
            if (endpart > 19) endpart = (byte)(endpart % 10);
            switch (endpart)
            {
                case 1: str += word1; break;
                case 2:
                case 3:
                case 4: str += word234; break;
                default: str += wordmore; break;
            }
            byte fracpart = decimal.ToByte((money - decintpart) * 100M);
            str += " " + ((fracpart < 10) ? "0" : "") + fracpart.ToString() + " ";
            if (fracpart > 19) fracpart = (byte)(fracpart % 10);
            switch (fracpart)
            {
                case 1: str += sword1; break;
                case 2:
                case 3:
                case 4: str += sword234; break;
                default: str += swordmore; break;
            };
            return str;
        }
        private static string EndTh(byte ThNum, byte Dek)
        {
            bool In234 = ((Dek >= 2) && (Dek <= 4));
            bool More4 = ((Dek > 4) || (Dek == 0));
            if (((ThNum > 2) && In234) || ((ThNum == 2) && (Dek == 1))) return "а";
            else if ((ThNum > 2) && More4) return "ов";
            else if ((ThNum == 2) && In234) return "и";
            else return "";
        }
        private static string EndDek1(byte Dek, bool IsMale)
        {
            if ((Dek > 2) || (Dek == 0)) return "";
            else if (Dek == 1)
            {
                if (IsMale) return "ин";
                else return "на";
            }
            else
            {
                if (IsMale) return "а";
                else return "е";
            }
        }
        #endregion
    }

    public class excel
    {
        public static string path;

        public static Excel.Application excelapp;

        public static Excel.Workbooks excelappworkbooks;

        public static Excel.Workbook excelappworkbook;

        public static Excel.Sheets excelsheets;

        public static Excel.Worksheet excelworksheet1, excelworksheet2;

        static excel()
        {
            excelapp = new Excel.Application();

            Microsoft.Win32.RegistryKey excelpath = null;
            try
            {
                excelpath = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Bobcalc");
                path = Convert.ToString(excelpath.GetValue("ExcelPath"));
            }
            finally
            {
                if (excelpath != null) excelpath.Close();
            }

            if (path != "")
            {
                excelapp.Visible = false;
                //Открываем книгу и получаем на нее ссылку
                excelappworkbook = excelapp.Workbooks.Open(path + "Bobmaster Calculator");
                //Получаем массив ссылок на листы выбранной книги
                excelsheets = excelappworkbook.Worksheets;
                //Получаем ссылку на лист 1 и 2              
                excelworksheet1 = (Excel.Worksheet)excelsheets.get_Item(1);
                excelworksheet2 = (Excel.Worksheet)excelsheets.get_Item(2);

            }
        }

        public static void changevisible(bool value)
        {
            excelapp.Visible = value;
        }
        //добавить сохранение пдф
        //в последней форме при нажатии на кнопку завершить - появляется окно(Желаете просмотреть КП перед версткой ПДФ файла?) 
        //и там условие мол если да-то change visible(true) а если нет то тупо сохранение пдф
    }

    static class Program
    {   public static void Senddata(string name, string unit, int count, int price, double discount, Excel.Application excelapp, Excel.Worksheet excelworksheet)
        { 
            int k = 10;
            int l;
            //Поиск пустой строки
             for (k = 10; k < 200; k++)
                {                  
                    Excel.Range cell = (Excel.Range)excelworksheet.Cells[k, 3];
                if (cell.Value == null ) { break; }
                }

            //Очистка на всякий пожарный, перед записью
            for (l = k; l < k + 6; l++)
            {
                for (int i = 2; i < 9; i++)
                {
                    Excel.Range cell = (Excel.Range)excel.excelworksheet1.Cells[l, i];
                    cell.Value = null;
                }
            }
            Excel.Range cell2 = (Excel.Range)excelworksheet.Cells[k + 1, 2];
            cell2.Value = "";
            cell2 = (Excel.Range)excelworksheet.Cells[k + 1, 3];
            cell2.Value = "";
            cell2 = (Excel.Range)excelworksheet.Cells[k + 2, 2];
            cell2.Value = "";
            cell2 = (Excel.Range)excelworksheet.Cells[k + 2, 3];
            cell2.Value = "";
            //
           
            //Заполнение строки и динамическая отрисовка таблицы
            cell2 = (Excel.Range)excelworksheet.Cells[k, 3];
            cell2.Value = name;
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlLeft;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;
            //Обводка
            cell2.Borders.ColorIndex = 0;            
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = (Excel.Range)excelworksheet.Cells[k, 4];
            cell2.Value = unit;
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;
            //Обводка
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = (Excel.Range)excelworksheet.Cells[k, 5];
            cell2.Value = count;
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;
            //Обводка
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = (Excel.Range)excelworksheet.Cells[k, 6];
            cell2.Value = price;
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;
            //Обводка
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = (Excel.Range)excelworksheet.Cells[k, 8];
            cell2.FormulaLocal = "=G" + Convert.ToString(k) + "*" + Convert.ToString(discount/100);

            cell2 = excelworksheet.get_Range("F" + k);
            cell2.Font.Bold = false;

            cell2 = (Excel.Range)excelworksheet.Cells[k, 7];
          //cell2.Value = count * price;
            cell2.FormulaLocal = "=E" + Convert.ToString(k) + "*F"+ Convert.ToString(k);
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;
            //Обводка
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = excelworksheet.get_Range("G" + k);
            cell2.Font.Bold = false;

            // столбец №
            cell2 = excelworksheet.get_Range("B11:B" + k);
            cell2.Merge();
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;
            cell2 = excelworksheet.get_Range("B10", "G" + (k));
            cell2.Font.Bold = false;
            //
            k = k + 1;

            cell2 = excelworksheet.get_Range("F" + k);
            cell2.Value = "Итого:";
            cell2.Font.Bold = true;

            cell2 = excelworksheet.get_Range("G" + k);
            cell2.FormulaLocal = "=СУММ(G11:G" + (k - 1) + ")";
            cell2.Font.Bold = true;

            cell2 = excelworksheet.get_Range("F" + (k + 1));
            cell2.Value = "Сумма скидки:";
            cell2.Font.Bold = true;

            cell2 = excelworksheet.get_Range("G" + (k + 1));
            cell2.FormulaLocal = "=СУММ(H11:H" + (k - 1) + ")";
            cell2.Font.Bold = true;

            cell2 = excelworksheet.get_Range("F" + (k + 2));
            cell2.Value = "Итого со скидкой:";
            cell2.Font.Bold = true;

            cell2 = excelworksheet.get_Range("G" + (k + 2));
            cell2.FormulaLocal = "=G" + k + "-G" + (k + 1);
            cell2.Font.Bold = true;

            //Итого прописью:
            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 2];
            cell2.Value = "Итого: ";
            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 3];
            Excel.Range cell3 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 7];
            cell2.Value = NumByWords.RurPhrase(Convert.ToDecimal(cell3.Value));


            /* Excel.Workbooks excelappworkbooks = excelapp.Workbooks;
             Excel.Workbook excelappworkbook = excelappworkbooks["Bobmaster Calculator"];
             excelappworkbook.Saved = false;*/
            //excelapp.Windows[1].Close(true, "E:\\Bobmaster\\Bobmaster Calculator");


        }

        public static void Send_GW_data(string OutGW, string OutEx, int price, double discount)
        {
            int k;
            int l;
            //Поиск пустой строки
            for (k = 10; k < 200; k++)
            {
                Excel.Range cell = (Excel.Range)excel.excelworksheet1.Cells[k, 3];
                if (cell.Value == null) { break; }
            }

            //Очистка на всякий пожарный, перед записью
            for (l = k; l < k + 6; l++)
            {
                for (int i = 2; i < 9; i++)
                {
                    Excel.Range cell = (Excel.Range)excel.excelworksheet1.Cells[l, i];
                    cell.Value = null;
                }
            }
            Excel.Range cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 1, 2];
            cell2.Value = "";
            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 1, 3];
            cell2.Value = "";
            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 2];
            cell2.Value = "";
            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 3];
            cell2.Value = "";


            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k, 3];
            cell2.Value = OutGW;
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlLeft;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;

            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k, 4];
            cell2.Value = "шт.";
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;

            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k, 5];
            cell2.Value = 1;
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;

            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k, 7];
            cell2.FormulaLocal = "=E" + Convert.ToString(k) + "*F" + Convert.ToString(k);
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;

            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k, 8];
            cell2.FormulaLocal = "=G" + Convert.ToString(k) + "*" + Convert.ToString(discount/100);

            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k, 6];
            cell2.Value = price;
            //Ориентация
            cell2.HorizontalAlignment = Excel.Constants.xlCenter;
            cell2.VerticalAlignment = Excel.Constants.xlCenter;
            cell2.WrapText = true;

            k = k + 1;

            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k, 3];
            cell2.Value = OutEx;
            cell2 = excel.excelworksheet1.get_Range("C" + (k - 1) + ":C" + k);
            cell2.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, 0);

            cell2 = excel.excelworksheet1.get_Range("D" + (k - 1) + ":D" + k);
            cell2.Merge();
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = excel.excelworksheet1.get_Range("E" + (k - 1) + ":E" + k);
            cell2.Merge();
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = excel.excelworksheet1.get_Range("F" + (k - 1) + ":F" + k);
            cell2.Merge();
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = excel.excelworksheet1.get_Range("G" + (k - 1) + ":G" + k);
            cell2.Merge();
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            cell2 = excel.excelworksheet1.get_Range("G" + k);
            cell2.Font.Bold = false;

            // столбец №
            cell2 = excel.excelworksheet1.get_Range("B11:B" + k);
            cell2.Merge();
            cell2.Borders.ColorIndex = 0;
            cell2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell2.Borders.Weight = Excel.XlBorderWeight.xlThin;
            cell2 = excel.excelworksheet1.get_Range("B10", "G" + (k));
            cell2.Font.Bold = false;
            //
            k = k + 1;

            cell2 = excel.excelworksheet1.get_Range("F" + k);
            cell2.Value = "Итого:";
            cell2.Font.Bold = true;

            cell2 = excel.excelworksheet1.get_Range("G" + k);
            cell2.FormulaLocal = "=СУММ(G11:G" + (k - 1) + ")";
            cell2.Font.Bold = true;

            cell2 = excel.excelworksheet1.get_Range("F" + (k + 1));
            cell2.Value = "Сумма скидки:";
            cell2.Font.Bold = true;

            cell2 = excel.excelworksheet1.get_Range("G" + (k + 1));
            cell2.FormulaLocal = "=СУММ(H11:H" + (k - 1) + ")";
            cell2.Font.Bold = true;

            cell2 = excel.excelworksheet1.get_Range("F" + (k + 2));
            cell2.Value = "Итого со скидкой:";
            cell2.Font.Bold = true;

            cell2 = excel.excelworksheet1.get_Range("G" + (k + 2));
            cell2.FormulaLocal = "=G" + k + "-G" + (k + 1);
            cell2.Font.Bold = true;

            //Итого прописью:
            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 2];
            cell2.Value = "Итого: ";
            cell2 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 3];
            Excel.Range cell3 = (Excel.Range)excel.excelworksheet1.Cells[k + 2, 7];
            cell2.Value = NumByWords.RurPhrase(Convert.ToDecimal(cell3.Value));
        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            System.Windows.Forms.Button buttonOk = new System.Windows.Forms.Button();
            System.Windows.Forms.Button buttonCancel = new System.Windows.Forms.Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        

        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new MainPage());

        }
    }
}
