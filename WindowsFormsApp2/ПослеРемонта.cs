﻿using MaterialSkin.Controls;
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
    public partial class ПослеРемонта : MaterialForm
    {

        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        string Код_СИ;
        public ПослеРемонта()
        {
            InitializeComponent();
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
        }

        private void ПослеРемонта_Load(object sender, EventArgs e)
        {

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

            SqlCommand cmd2 = new SqlCommand("select count(*) from Ремонт", conn);
            conn.Open();
            int count = (int)cmd2.ExecuteScalar();
            conn.Close();

            SqlCommand cmd3 = new SqlCommand("insert into Ремонт(Код_ремонта, Код_СИ, Дата_поступления, Дата_завершения,Неисправность, Стоимость) values (@Код_ремонта, @Код_СИ, @Дата_поступления, @Дата_завершения,@Неисправность, @Стоимость)", conn);

            conn.Open();
            cmd3.Parameters.AddWithValue("@Код_ремонта", SqlDbType.Int).Value = count + 1;
            cmd3.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
            cmd3.Parameters.AddWithValue("@Дата_поступления", dateTimePicker1.Text);
            cmd3.Parameters.AddWithValue("@Дата_завершения", dateTimePicker2.Text);
            cmd3.Parameters.AddWithValue("@Неисправность", textBox1.Text);
            cmd3.Parameters.AddWithValue("@Стоимость", textBox2.Text);


            try
            {
                cmd3.ExecuteNonQuery();
                MessageBox.Show("Инструмент добавлен!");
            }
            catch
            { MessageBox.Show("Не удалось добавить инструмент!"); }
            finally { conn.Close(); }
        }
    }
}
