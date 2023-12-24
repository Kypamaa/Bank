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
using System.Data.SQLite;

namespace Bank
{
    public partial class Form_Authorization : Form
    {
        private SQLiteConnection sqlite;

        string login;
        string fullname;
        string signature;

        public Form_Authorization()
        {
            InitializeComponent();
            sqlite = new SQLiteConnection("Data Source=database.db; Version=3;");
            sqlite.Open();
        }

        private string[] GetUser(string login)
        {
            string query = $"SELECT * FROM Users WHERE Login='{login}'";
            SQLiteCommand command = new SQLiteCommand(query, sqlite);
            SQLiteDataReader reader = command.ExecuteReader();

            if (reader.Read())
            {
                return new string[] {
                    reader["Login"].ToString(),
                    reader["FullName"].ToString(),
                    reader["Password"].ToString(),
                    reader["Signature"].ToString()
                };
            }

            return null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Предполагается, что textBoxLogin и textBoxPassword - это TextBox, из которых вы хотите получить значения
            string[] user = GetUser(textBox1.Text);

            if (user != null && user[2] == textBox2.Text)
            {
                login = user[0].ToString();
                fullname = user[1].ToString();
                signature = user[3].ToString();

                Form_Main form_main = new Form_Main(login, fullname, signature);
                this.Hide();
                form_main.Show();
            }
            else
            {
                label5.Visible = true;
                textBox1.Clear();
                textBox2.Clear();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form_Registration form_reg = new Form_Registration();
            this.Hide();
            form_reg.Show();
        }
    }
}