using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Data.Entity;

namespace Bank
{
    public partial class Form_Registration : Form
    {
        private SQLiteConnection sqlite;

        public Form_Registration()
        {
            InitializeComponent();
            CreateDatabase();
        }

        private void CreateDatabase()
        {
            // Создание новой базы данных, если она еще не существует
            sqlite = new SQLiteConnection("Data Source=database.db; Version=3;");
            sqlite.Open();

            // Создание таблицы
            string query = "CREATE TABLE IF NOT EXISTS Users (Login VARCHAR(20), FullName VARCHAR(50), Password VARCHAR(50), Signature VARCHAR(50))";
            SQLiteCommand command = new SQLiteCommand(query, sqlite);
            command.ExecuteNonQuery();
        }

        private void AddUser(string login, string fullName, string password, string signature)
        {
            Form_Authorization form_auth = new Form_Authorization();

            string query = $"INSERT INTO Users (Login, FullName, Password, Signature) VALUES ('{login}', '{fullName}', '{password}', '{signature}')";
            SQLiteCommand command = new SQLiteCommand(query, sqlite);
            command.ExecuteNonQuery();

            this.Close();
            form_auth.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Предполагается, что textBoxLogin, textBoxFullName, textBoxPassword и textBoxSignature - это TextBox, из которых вы хотите получить значения
            AddUser(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form_Authorization form_auth = new Form_Authorization();

            this.Close();
            form_auth.Show();
        }
    }
}