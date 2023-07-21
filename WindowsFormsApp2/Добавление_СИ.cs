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
    public partial class Добавление_СИ : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        public Добавление_СИ()
        {
            InitializeComponent();
        }

        private void Добавление_СИ_Load(object sender, EventArgs e)
        {

            // Заполнение комбобокса 1
            SqlConnection conn = new SqlConnection(connection);
            string command = "select Наименование from Наименование_СИ";
            SqlCommand cmd = new SqlCommand(command, conn);
            conn.Open();
            SqlDataReader DR = cmd.ExecuteReader();

            while (DR.Read())
            {
                comboBox5.Items.Add(DR[0]);

            }
            conn.Close();

            // Заполнение комбобокса 2
           
            string command2 = "select Тип from Тип_инструмента";
            SqlCommand cmd2 = new SqlCommand(command2, conn);
            conn.Open();
            SqlDataReader DR2 = cmd2.ExecuteReader();

            while (DR2.Read())
            {
                comboBox4.Items.Add(DR2[0]);

            }
            conn.Close();

            // Заполнение комбобокса 3
            
            string command3 = "select Наименование from Единица_измерения";
            SqlCommand cmd3 = new SqlCommand(command3, conn);
            conn.Open();
            SqlDataReader DR3 = cmd3.ExecuteReader();

            while (DR3.Read())
            {
                comboBox3.Items.Add(DR3[0]);

            }
            conn.Close();

            // Заполнение комбобокса 4
          
            string command4 = "select Страна from Страна_изготовитель";
            SqlCommand cmd4 = new SqlCommand(command4, conn);
            conn.Open();
            SqlDataReader DR4 = cmd4.ExecuteReader();

            while (DR4.Read())
            {
                comboBox1.Items.Add(DR4[0]);

            }
            conn.Close();

            // Заполнение комбобокса 5
         
            string command5 = "select Год from Год_выпуска";
            SqlCommand cmd5 = new SqlCommand(command5, conn);
            conn.Open();
            SqlDataReader DR5 = cmd5.ExecuteReader();

            while (DR5.Read())
            {
                comboBox2.Items.Add(DR5[0]);

            }
            conn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            SqlConnection conn = new SqlConnection(connection);

            SqlCommand cmd2 = new SqlCommand("select count(*) from Средство_измерения", conn);
            conn.Open();
            int count = (int)cmd2.ExecuteScalar();
            conn.Close();

            SqlCommand cmd = new SqlCommand("insert into Средство_измерения (Код_СИ, Код_наименования, Код_типа, Заводской_номер, Код_единицы_измерения, Код_страны, Код_года)" +
                "values (@Код_СИ, @Код_наименования, @Код_типа, @Заводской_номер, @Код_единицы_измерения, @Код_страны, @Код_года) ", conn);
            
            conn.Open();

            int Код_наименования = Convert.ToInt32(comboBox5.SelectedIndex);
            int Код_типа = Convert.ToInt32(comboBox4.SelectedIndex);
            int Код_единицы_измерения = Convert.ToInt32(comboBox3.SelectedIndex);
            int Код_страны = Convert.ToInt32(comboBox1.SelectedIndex);
            int Код_года = Convert.ToInt32(comboBox2.SelectedIndex);

            cmd.Parameters.Add("@Код_СИ", SqlDbType.Int).Value = count + 1;
            cmd.Parameters.Add("@Код_наименования", SqlDbType.Int).Value = Код_наименования + 1;
            cmd.Parameters.Add("@Код_типа", SqlDbType.Int).Value = Код_типа + 1;
            cmd.Parameters.Add("@Код_единицы_измерения", SqlDbType.Int).Value = Код_единицы_измерения + 1;
            cmd.Parameters.Add("@Код_страны", SqlDbType.Int).Value = Код_страны + 1;
            cmd.Parameters.Add("@Код_года", SqlDbType.Int).Value = Код_года + 1;
            cmd.Parameters.Add("@Заводской_номер", textBox1.Text);

            try
            {
                cmd.ExecuteNonQuery();
                MessageBox.Show("Инструмент внесен!");
            }
            catch
            { MessageBox.Show("Не удалось совершить внесение инструмента!"); }
            finally { conn.Close(); }
        }
    }
}
