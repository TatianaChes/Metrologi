using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Добавить_в_план : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        string Код_СИ;
        public Добавить_в_план()
        {
            InitializeComponent();
        }

        private void Добавить_в_план_Load(object sender, EventArgs e)
        {
            // Заполнение комбобокса 
            SqlConnection conn = new SqlConnection(connection);
            string command = "select Заводской_номер from Средство_измерения";
            SqlCommand cmd = new SqlCommand(command, conn);
            conn.Open();
            SqlDataReader DR = cmd.ExecuteReader();

            while (DR.Read())
            {
                comboBox1.Items.Add(DR[0]);

            }
            conn.Close();


            // Заполнение комбобокса 2
            string command2 = "select Наименование from Вид_метрологического_контроля";
            SqlCommand cmd2 = new SqlCommand(command2, conn);
            conn.Open();
            SqlDataReader DR2 = cmd2.ExecuteReader();

            while (DR2.Read())
            {
                comboBox2.Items.Add(DR2[0]);

            }
            conn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlCommand cmd = new SqlCommand("select Код_СИ from Средство_измерения where Заводской_номер ='" + comboBox1.Text + "'", conn);
            conn.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Код_СИ = reader["Код_СИ"].ToString();

            }
            conn.Close();

            SqlCommand cmd2 = new SqlCommand("select count(*) from План_контроля", conn);
            conn.Open();
            int count = (int)cmd2.ExecuteScalar();
            conn.Close();

            SqlCommand cmd3 = new SqlCommand("insert into План_контроля(Код_плана, Код_СИ, Код_вида, МПИ, Дата_следующей_поверки) values (@Код_плана, @Код_СИ, @Код_вида, @МПИ, @Дата_следующей_поверки)", conn);
            int Код_вида = Convert.ToInt32(comboBox2.SelectedIndex);
            conn.Open();
            cmd3.Parameters.AddWithValue("@Код_плана", SqlDbType.Int).Value = count + 1;
            cmd3.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
            cmd3.Parameters.AddWithValue("@Код_вида", SqlDbType.Int).Value = Код_вида + 1;
            cmd3.Parameters.AddWithValue("@МПИ", textBox1.Text);
            cmd3.Parameters.AddWithValue("@Дата_следующей_поверки", textBox2.Text);

            try
            {
                cmd3.ExecuteNonQuery();
                MessageBox.Show("Инструмент добавлен в план!");
            }
            catch
            { MessageBox.Show("Не удалось добавить инструмент!"); }
            finally { conn.Close(); }
        }
    }
}
