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
using MaterialSkin;
using MaterialSkin.Controls;
namespace WindowsFormsApp2
{
    public partial class EventForm : MaterialForm
    {
      
        string connection = @"Data Source=.\SQLEXPRESS;Initial Catalog=Метрология;Integrated Security=True";

        public EventForm()
        {
            InitializeComponent();
        }

        private void EventForm_Load(object sender, EventArgs e)
        {
            textBox1.Text = Form3.static_year + "/" + Form3.static_month + "/" +UserControlDays.static_day;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlCommand command = new SqlCommand("insert into Событие(Дата, Событие) values (@Дата,@Событие)", conn);
            conn.Open();

            command.Parameters.AddWithValue("@Дата", textBox1.Text);
            command.Parameters.AddWithValue("@Событие", textBox2.Text);


            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Событие создано!");
            }
            catch
            { MessageBox.Show("Не удалось создать событие!"); }
            finally { conn.Close(); }
        }
    }
}
