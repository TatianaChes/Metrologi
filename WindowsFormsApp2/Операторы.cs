using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Операторы : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология;Integrated Security = True";
        int selectedRow;
        string password;
        public Операторы()
        {
            InitializeComponent();
        }

        private void Операторы_Load(object sender, EventArgs e)
        {
            Print();


        }

        private void Print()
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter("select * from Оператор", conn);
            DataTable table = new DataTable();
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            adapter.Fill(table);
            dataGridView1.DataSource = table;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRow = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[selectedRow];

            textBox1.Text = row.Cells[0].Value.ToString();
            textBox2.Text = row.Cells[1].Value.ToString();
            password = row.Cells[1].Value.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы действительно хотите удалить запись?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
            {

            }
            else
            {

                try
                {
                    SqlCommand cmd;
                    SqlConnection con;
                    con = new SqlConnection(connection);
                    cmd = new SqlCommand("DELETE FROM Оператор WHERE Логин = @Логин", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@Логин", dataGridView1[0, selectedRow].Value.ToString());
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Оператор удален!");
                    Print();
                }
                catch
                {
                    MessageBox.Show("Не удалось удалить");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlCommand command = new SqlCommand("insert into Оператор (Логин, Пароль) values(@Логин, @Пароль)", conn);
            conn.Open();
            command.Parameters.AddWithValue("@Логин", textBox1.Text);
            command.Parameters.AddWithValue("@Пароль", GetHash(textBox2.Text));
            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Оператор создан!");
            }
            catch
            { MessageBox.Show("Не удалось создать оператора!"); }
            finally { conn.Close(); Print(); }


        }

        public string GetHash(string input)
        {
            var md5 = MD5.Create();
            var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));

            return Convert.ToBase64String(hash);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlCommand command = new SqlCommand("update Оператор set Логин = @Логин  , Пароль =  @Пароль where Логин = '" + dataGridView1[0, selectedRow].Value.ToString()+"'", conn);
            conn.Open();
            command.Parameters.AddWithValue("@Логин", textBox1.Text);

            if (checkBox1.Checked == true)
            {
                command.Parameters.AddWithValue("@Пароль", GetHash(textBox2.Text));
            }
            else
            {
                command.Parameters.AddWithValue("@Пароль", password);
            }
           
            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Оператор обновлен!");
            }
            catch
            { MessageBox.Show("Не удалось обновить оператора!"); }
            finally { conn.Close(); Print(); }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.ReadOnly = false;
        }
    }
}
