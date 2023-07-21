using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Производство : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        string changedCell;
        string changedCell2;
        int selectedRow;
        public Производство()
        {
            InitializeComponent();
        }

        private void Производство_Load(object sender, EventArgs e)
        {
            Print();
        }

        private void Print()
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter("select Код_инструмента_в_работе as Код, Наименование_СИ.Наименование, " +
                "Тип_инструмента.Тип, Средство_измерения.Заводской_номер,Дата_ввода from Инструмент_в_работе join Средство_измерения on" +
                " Средство_измерения.Код_СИ = Инструмент_в_работе.Код_СИ join Наименование_СИ on Наименование_СИ.Код_наименования =" +
                " Средство_измерения.Код_наименования left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа", conn);
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;

            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[1].Width = 110;
            dataGridView1.Columns[3].Width = 150;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Добавить_в_производство в_производство = new Добавить_в_производство();
            в_производство.Show();
            в_производство.FormClosed += new FormClosedEventHandler(af_FormClosed);
        }

        private void af_FormClosed(object sender, FormClosedEventArgs e)
        {
            Print();
        }

        private void button6_Click(object sender, EventArgs e) 
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
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.LightGray;
                            break;
                        }
                
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter("select Код_инструмента_в_работе as Код, Наименование_СИ.Наименование, " +
                "Тип_инструмента.Тип, Средство_измерения.Заводской_номер,Дата_ввода from Инструмент_в_работе join Средство_измерения on" +
                " Средство_измерения.Код_СИ = Инструмент_в_работе.Код_СИ join Наименование_СИ on Наименование_СИ.Код_наименования =" +
                " Средство_измерения.Код_наименования left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа" +
                " where Дата_ввода  between '"+ dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and '"+ dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' ", conn);
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;

            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[1].Width = 110;
            dataGridView1.Columns[3].Width = 150;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }
     
        private void printDocument1_PrintPage_1(object sender, PrintPageEventArgs e)
        {
            Bitmap bmp = new Bitmap(dataGridView1.Size.Width, dataGridView1.Size.Height);
            dataGridView1.DrawToBitmap(bmp, dataGridView1.Bounds);
            e.Graphics.DrawImage(bmp, 0, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            conn.Open();
            SqlCommand command = new SqlCommand("update Инструмент_в_работе set Дата_ввода = @Дата_ввода where Код_инструмента_в_работе=" + changedCell,conn);
            command.Parameters.AddWithValue("@Дата_ввода", changedCell2);
            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Данные обновлены!");
            }
            catch
            { MessageBox.Show("Не удалось обновить данные!"); }
            finally { conn.Close(); }
            Print();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
          changedCell = dataGridView1[0, e.RowIndex].Value.ToString();
            changedCell2 = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
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
                    cmd = new SqlCommand("DELETE FROM Инструмент_в_работе WHERE Код_инструмента_в_работе = @Код_инструмента_в_работе", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@Код_инструмента_в_работе", dataGridView1[0, selectedRow].Value.ToString());
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRow = e.RowIndex;
        }
    }
}
