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
    public partial class СИ : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        int selectedRow;
        string Код_наименования, Код_типа, Код_единицы_измерения, Код_страны, Код_года;
        public СИ()
        {
            InitializeComponent();
        }

        private void СИ_Load(object sender, EventArgs e)
        {
            Print();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRow = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[selectedRow];

            textBox1.Text = row.Cells[1].Value.ToString();
            textBox2.Text = row.Cells[2].Value.ToString();
            textBox3.Text = row.Cells[3].Value.ToString();
            textBox4.Text = row.Cells[4].Value.ToString();
            textBox5.Text = row.Cells[5].Value.ToString();
            textBox6.Text = row.Cells[6].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Добавление_СИ добавление = new Добавление_СИ();
            добавление.Show();
           добавление.FormClosed += new FormClosedEventHandler(af_FormClosed);
        }

        private void af_FormClosed(object sender, FormClosedEventArgs e)
        {
            Print();
        }

        private void Print()
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter(" select Код_СИ as Код, Наименование_СИ.Наименование, Тип_инструмента.Тип, Средство_измерения.Заводской_номер, Единица_измерения.Наименование as Ед_измерения, Страна_изготовитель.Страна, Год_выпуска.Год from Средство_измерения join Наименование_СИ on Наименование_СИ.Код_наименования = Средство_измерения.Код_наименования left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа left join Единица_измерения on Единица_измерения.Код_единицы_измерения = Средство_измерения.Код_единицы_измерения join Страна_изготовитель on Страна_изготовитель.Код_страны = Средство_измерения.Код_страны join Год_выпуска on Год_выпуска.Код_года = Средство_измерения.Код_года", conn);
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;


            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[3].Width = 110;
        }

        private void button3_Click(object sender, EventArgs e)
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
                    cmd = new SqlCommand("DELETE FROM Средство_измерения WHERE Код_СИ = @Код_СИ", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@Код_СИ", dataGridView1[0, selectedRow].Value.ToString());
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Инструмент удален!");
                    Print();
                }
                catch
                {
                    MessageBox.Show("Не удалось удалить");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.White;
                }

            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox7.Text))
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.LightGray;
                            break;
                        }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {

        
            SqlConnection conn = new SqlConnection(connection);
            SqlCommand cmd = new SqlCommand(" select Наименование_СИ.Код_наименования from Наименование_СИ where Наименование_СИ.Наименование = '"+ textBox1.Text+"'", conn);
            conn.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Код_наименования = reader["Код_наименования"].ToString();

            }
            conn.Close();

            SqlCommand cmd2 = new SqlCommand(" select Тип_инструмента.Код_типа from Тип_инструмента where Тип_инструмента.Тип = '" + textBox2.Text + "'", conn);
            conn.Open();
            SqlDataReader reader2 = cmd2.ExecuteReader();
            if (reader2.Read())
            {
                Код_типа = reader2["Код_типа"].ToString();

            }
            conn.Close();

            SqlCommand cmd3 = new SqlCommand(" select Код_единицы_измерения from Единица_измерения where Наименование = '" + textBox4.Text + "'", conn);
            conn.Open();
            SqlDataReader reader3 = cmd3.ExecuteReader();
            if (reader3.Read())
            {
                Код_единицы_измерения = reader3["Код_единицы_измерения"].ToString();

            }
            conn.Close();

            SqlCommand cmd4 = new SqlCommand(" select Код_страны from Страна_изготовитель where Страна = '" + textBox5.Text + "'", conn);
            conn.Open();
            SqlDataReader reader4 = cmd4.ExecuteReader();
            if (reader4.Read())
            {
                Код_страны = reader4["Код_страны"].ToString();

            }
            conn.Close();

            SqlCommand cmd5 = new SqlCommand(" select Код_года from Год_выпуска where Год =" + textBox6.Text, conn);
            conn.Open();
            SqlDataReader reader5 = cmd5.ExecuteReader();
            if (reader5.Read())
            {
                Код_года = reader5["Код_года"].ToString();

            }
            conn.Close();

            SqlCommand command = new SqlCommand("update Средство_измерения set Код_наименования = @Код_наименования, Код_типа = @Код_типа, Заводской_номер = @Заводской_номер, Код_единицы_измерения = @Код_единицы_измерения, Код_страны = @Код_страны, Код_года = @Код_года where Код_СИ =" + dataGridView1[0, selectedRow].Value.ToString(), conn);
            conn.Open();
            command.Parameters.AddWithValue("@Код_наименования", SqlDbType.Int).Value = Convert.ToInt32(Код_наименования);
            command.Parameters.AddWithValue("@Код_типа", SqlDbType.Int).Value = Convert.ToInt32(Код_типа);
            command.Parameters.AddWithValue("@Код_единицы_измерения", SqlDbType.Int).Value = Convert.ToInt32(Код_единицы_измерения);
            command.Parameters.AddWithValue("@Код_страны", SqlDbType.Int).Value = Convert.ToInt32(Код_страны);
            command.Parameters.AddWithValue("@Код_года", SqlDbType.Int).Value = Convert.ToInt32(Код_года);
            command.Parameters.AddWithValue("@Заводской_номер", textBox3.Text);

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Данные обновлены!");
            }
            catch
            { MessageBox.Show("Не удалось обновить данные!"); }
            finally { conn.Close(); Print();}


        }
    }
}
