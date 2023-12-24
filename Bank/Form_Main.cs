using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Data.SQLite;

namespace Bank
{
    public partial class Form_Main : Form
    {
        private SQLiteConnection sqlite;

        float amount;
        int term;
        int replenishment;

        double result1_panel;
        double result2_panel;
        double result3_panel;

        double final_balance1;
        double final_balance2;
        double final_balance3;

        private string login;
        private string fullname;
        private string signature;

        public Form_Main(string login, string fullname, string signature)
        {
            InitializeComponent();
            panel12.BringToFront();

            sqlite = new SQLiteConnection("Data Source=database.db; Version=3;");
            sqlite.Open();

            this.login = login;
            this.fullname = fullname;
            this.signature = signature;
        }
        private async void panel12_MouseClick(object sender, MouseEventArgs e)
        {

            while (panel12.Location.Y != panel10.Location.Y)
            {
                await System.Threading.Tasks.Task.Delay(1);
                panel12.Location = new System.Drawing.Point(1, panel12.Location.Y - 10);
            }
        }
        private async void button1_Click(object sender, EventArgs e)
        {
            while (panel12.Location.Y != 580)
            {
                await System.Threading.Tasks.Task.Delay(1);
                panel12.Location = new System.Drawing.Point(0, panel12.Location.Y + 10);
            }
        }
        private async void button2_Click(object sender, EventArgs e)
        {
            label46.Text = result1_panel.ToString("0,0");
            label47.Text = result2_panel.ToString("0,0");
            label53.Text = result3_panel.ToString("0,0");

            label49.Text = final_balance1.ToString("0,0");
            label45.Text = final_balance2.ToString("0,0");
            label51.Text = final_balance3.ToString("0,0");

            while (panel14.Location.Y != panel13.Location.Y)
            {
                await System.Threading.Tasks.Task.Delay(1);
                panel14.Location = new System.Drawing.Point(1, panel14.Location.Y - 10);

                await System.Threading.Tasks.Task.Delay(1);
                panel12.Location = new System.Drawing.Point(0, panel12.Location.Y + 1);
            }
        }

        public static (double, double) Result1(double balance, int months)
        {
            double monthlyRate = 0.08 / 12;
            double totalIncome = 0;

            for (int i = 0; i < months; i++)
            {
                double income = balance * monthlyRate;
                totalIncome += income;
                balance += income;
            }

            return (totalIncome, balance);
        }
        public static (double, double) Result2(double balance, int months, int deposit)
        {
            double monthlyRate = 0.05 / 12;
            double totalIncome = 0;

            for (int i = 0; i < months; i++)
            {
                balance += deposit; // Пополнение счета
                double income = balance * monthlyRate;
                totalIncome += income;
                balance += income; // Капитализация процентов
            }

            return (totalIncome, balance);
        }
        public static (double, double) Result3(double balance, int months, int deposit)
        {
            double totalIncome = 0;

            for (int i = 0; i < months; i++)
            {
                balance += deposit;
                totalIncome += balance * (0.06 / 12);
            }

            return (totalIncome, balance);
        }
        private void UpdateResult()
        {
            (double result1, double balance1) = Result1(amount, term);
            (double result2, double balance2) = Result2(amount, term, replenishment);
            (double result3, double balance3) = Result3(amount, term, replenishment);

            final_balance1 = balance1;
            final_balance2 = balance2;
            final_balance3 = balance3;

            result1_panel = result1;
            result2_panel = result2;
            result3_panel = result3;

            result1 = Math.Truncate(result1);
            result2 = Math.Truncate(result2);
            result3 = Math.Truncate(result3);

            if (term < 1)
            {
                textBox1.Text = "Минимальный срок от 3 месяцев";
                textBox3.Text = "Минимальный срок от 3 месяцев";
            }
            else
            {
                textBox1.Text = result1.ToString();
                textBox3.Text = result3.ToString();
            }

            if (term < 3)
            {
                textBox2.Text = "Минимальный срок от 6 месяцев";
            }
            else
            {
                textBox2.Text = result2.ToString();
            }
        }


        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            textBox4.Text = trackBar1.Value.ToString();
            amount = trackBar1.Value;

            UpdateResult();
        }
        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            xz();
            UpdateResult();
        }
        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            textBox6.Text = trackBar3.Value.ToString();
            replenishment = trackBar3.Value;

            UpdateResult();
        }

        void xz()
        {
            textBox5.Text = (30 + 30 * trackBar2.Value).ToString();
            term = trackBar2.Value + 1;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PNG Image|*.png";
            saveFileDialog.Title = "Save PNG File";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                // Создание скриншота панели
                Bitmap bitmap = new Bitmap(999, 650);
                panel14.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, 999, 650));

                // Сохранение скриншота в PNG файл
                bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Png);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Здесь предполагается, что у вас есть текстовые поля с именами, соответствующими вашим переменным.
            // Замените "Variable1", "Variable2" и т.д. на имена ваших переменных.

            // Создайте новый экземпляр приложения Word и откройте документ.
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(@"C:\Users\Kypama\Desktop\Bank 2.5\Bank\Resources\Договор.docx");

            string Variable1 = fullname.ToString();
            string Variable2 = login.ToString();
            string Variable3 = amount.ToString();
            string Variable4 = textBox5.Text.ToString();
            string Variable5 = replenishment.ToString();
            string Variable6 = label49.Text.ToString();
            string Variable7 = label46.Text.ToString();
            string Variable8 = "Стабильный";
            string Variable9 = "8";
            string Variable10 = DateTime.Now.ToString();
            string Variable11 = signature.ToString();

            // Замените метки в документе на значения переменных.
            FindAndReplace(wordApp, "FIO", Variable1);
            FindAndReplace(wordApp, "LOGIN_VKLAD", Variable2);
            FindAndReplace(wordApp, "SUM_VKLAD", Variable3);
            FindAndReplace(wordApp, "DATE_OFF", Variable4);
            FindAndReplace(wordApp, "REPLA", Variable5);
            FindAndReplace(wordApp, "FINAL_SUM", Variable6);
            FindAndReplace(wordApp, "DOHOD", Variable7);
            FindAndReplace(wordApp, "VKLAD_NAME", Variable8);
            FindAndReplace(wordApp, "PROCENT", Variable9);
            FindAndReplace(wordApp, "DATE_ON", Variable10);
            FindAndReplace(wordApp, "PODPIC", Variable11);

            // Предложите пользователю выбрать место для сохранения нового файла.
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Document|*.docx";
            saveFileDialog.Title = "Сохранить документ Word";
            saveFileDialog.ShowDialog();

            // Если пользователь выбрал место для сохранения, сохраните документ там.
            if (saveFileDialog.FileName != "")
            {
                wordDoc.SaveAs2(saveFileDialog.FileName);
            }

            // Закройте документ.
            wordDoc.Close();
            wordApp.Quit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Здесь предполагается, что у вас есть текстовые поля с именами, соответствующими вашим переменным.
            // Замените "Variable1", "Variable2" и т.д. на имена ваших переменных.

            // Создайте новый экземпляр приложения Word и откройте документ.
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(@"C:\Users\Kypama\Desktop\Bank 2.5\Bank\Resources\Договор.docx");

            string Variable1 = fullname.ToString();
            string Variable2 = login.ToString();
            string Variable3 = amount.ToString();
            string Variable4 = textBox5.Text.ToString();
            string Variable5 = replenishment.ToString();
            string Variable6 = label45.Text.ToString();
            string Variable7 = label47.Text.ToString();
            string Variable8 = "Оптимальный";
            string Variable9 = "5";
            string Variable10 = DateTime.Now.ToString();
            string Variable11 = signature.ToString();

            // Замените метки в документе на значения переменных.
            FindAndReplace(wordApp, "FIO", Variable1);
            FindAndReplace(wordApp, "LOGIN_VKLAD", Variable2);
            FindAndReplace(wordApp, "SUM_VKLAD", Variable3);
            FindAndReplace(wordApp, "DATE_OFF", Variable4);
            FindAndReplace(wordApp, "REPLA", Variable5);
            FindAndReplace(wordApp, "FINAL_SUM", Variable6);
            FindAndReplace(wordApp, "DOHOD", Variable7);
            FindAndReplace(wordApp, "VKLAD_NAME", Variable8);
            FindAndReplace(wordApp, "PROCENT", Variable9);
            FindAndReplace(wordApp, "DATE_ON", Variable10);
            FindAndReplace(wordApp, "PODPIC", Variable11);

            // Предложите пользователю выбрать место для сохранения нового файла.
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Document|*.docx";
            saveFileDialog.Title = "Сохранить документ Word";
            saveFileDialog.ShowDialog();

            // Если пользователь выбрал место для сохранения, сохраните документ там.
            if (saveFileDialog.FileName != "")
            {
                wordDoc.SaveAs2(saveFileDialog.FileName);
            }

            // Закройте документ.
            wordDoc.Close();
            wordApp.Quit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Здесь предполагается, что у вас есть текстовые поля с именами, соответствующими вашим переменным.
            // Замените "Variable1", "Variable2" и т.д. на имена ваших переменных.

            // Создайте новый экземпляр приложения Word и откройте документ.
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(@"C:\Users\Kypama\Desktop\Bank 2.5\Bank\Resources\Договор.docx");

            string Variable1 = fullname.ToString();
            string Variable2 = login.ToString();
            string Variable3 = amount.ToString();
            string Variable4 = textBox5.Text.ToString();
            string Variable5 = replenishment.ToString();
            string Variable6 = label41.Text.ToString();
            string Variable7 = label53.Text.ToString();
            string Variable8 = "Стандарт";
            string Variable9 = "6";
            string Variable10 = DateTime.Now.ToString();
            string Variable11 = signature.ToString();

            // Замените метки в документе на значения переменных.
            FindAndReplace(wordApp, "FIO", Variable1);
            FindAndReplace(wordApp, "LOGIN_VKLAD", Variable2);
            FindAndReplace(wordApp, "SUM_VKLAD", Variable3);
            FindAndReplace(wordApp, "DATE_OFF", Variable4);
            FindAndReplace(wordApp, "REPLA", Variable5);
            FindAndReplace(wordApp, "FINAL_SUM", Variable6);
            FindAndReplace(wordApp, "DOHOD", Variable7);
            FindAndReplace(wordApp, "VKLAD_NAME", Variable8);
            FindAndReplace(wordApp, "PROCENT", Variable9);
            FindAndReplace(wordApp, "DATE_ON", Variable10);
            FindAndReplace(wordApp, "PODPIC", Variable11);

            // Предложите пользователю выбрать место для сохранения нового файла.
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Document|*.docx";
            saveFileDialog.Title = "Сохранить документ Word";
            saveFileDialog.ShowDialog();

            // Если пользователь выбрал место для сохранения, сохраните документ там.
            if (saveFileDialog.FileName != "")
            {
                wordDoc.SaveAs2(saveFileDialog.FileName);
            }

            // Закройте документ.
            wordDoc.Close();
            wordApp.Quit();
        }
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike,
                ref matchAllWordForms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiacritics, ref matchAlefHamza,
                ref matchControl);
        }

    }
}