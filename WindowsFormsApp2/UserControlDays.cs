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
    public partial class UserControlDays : UserControl
    {

        string connection = @"Data Source=.\SQLEXPRESS;Initial Catalog=Метрология;Integrated Security=True";

        public static string static_day;
        public UserControlDays()
        {
            InitializeComponent();
            timer1.Start();
        }

        private void UserControlDays_Load(object sender, EventArgs e)
        {
            
        }
        public void days(int numday) {
            label1.Text = numday + "";
        }

        private void UserControlDays_Click(object sender, EventArgs e)
        {
            static_day = label1.Text;
                     
            EventForm eventform = new EventForm();
            eventform.Show();

        }
        public void displayEvent() {
            SqlConnection conn = new SqlConnection(connection);
            conn.Open();
            SqlCommand command = new SqlCommand("select * from Событие where Дата = @Дата", conn);
            command.Parameters.AddWithValue("Дата", Form3.static_year + "/" + Form3.static_month + "/" + label1.Text);
            SqlDataReader reader = command.ExecuteReader();
            if (reader.Read()) {
                label2.Text = reader["Событие"].ToString();
            }
            reader.Dispose();
            command.Dispose();
            conn.Close();
        }


        public void displayEvent2()
        {
            SqlConnection conn = new SqlConnection(connection);
            conn.Open();
            SqlCommand command = new SqlCommand("select * from План_контроля join Средство_измерения on Средство_измерения.Код_СИ = План_контроля.Код_СИ where Дата_следующей_поверки = @Дата ", conn);
            command.Parameters.AddWithValue("Дата", Form3.static_year + "/" + Form3.static_month + "/" + label1.Text);
            SqlDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                label2.Text = reader["Заводской_номер"].ToString();
            }
            reader.Dispose();
            command.Dispose();
            conn.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            displayEvent();
            displayEvent2();
        }
    }
}
