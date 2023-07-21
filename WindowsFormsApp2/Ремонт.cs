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
    public partial class Ремонт : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        int selectedRow;
        string Код_СИ,changedCell;
        public Ремонт()
        {
            InitializeComponent();
        }

        private void Ремонт_Load(object sender, EventArgs e)
        {
            Print();
            label1.Visible = false;
            label2.Visible = false;
        }

        private void Print()
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter(" select Код_ремонта as Код, Наименование_СИ.Наименование, Тип_инструмента.Тип, " +
                "Средство_измерения.Заводской_номер,  Дата_поступления, Дата_завершения, Неисправность, Стоимость from Ремонт join " +
                "Средство_измерения on Средство_измерения.Код_СИ = Ремонт.Код_СИ join Наименование_СИ on Наименование_СИ.Код_наименования" +
                " = Средство_измерения.Код_наименования  left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа", conn);
            DataTable table = new DataTable();
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            adapter.Fill(table);
            dataGridView1.DataSource = table;

            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[4].Width = 110;
            dataGridView1.Columns[5].Width = 110;
            dataGridView1.Columns[6].Width = 130;
            dataGridView1.Columns[7].Width = 90;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ПослеРемонта после = new ПослеРемонта();
            после.Show();
            после.FormClosed += new FormClosedEventHandler(af_FormClosed);
        }

        private void af_FormClosed(object sender, FormClosedEventArgs e)
        {
            Print();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            Bitmap bm = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);
            dataGridView1.DrawToBitmap(bm, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));
            e.Graphics.DrawImage(bm, 0, 0);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label1.Visible = true;
            label2.Visible = true;
            SqlConnection conn = new SqlConnection(connection);
            conn.Open();
            SqlCommand command = new SqlCommand("select sum(Стоимость) as Стоимость from Ремонт", conn);
            SqlDataReader reader = command.ExecuteReader();
            if (reader.Read()) {
                label2.Text = reader["Стоимость"].ToString();
            }
            conn.Close();


        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlCommand cmd = new SqlCommand("select Код_СИ from Средство_измерения where Заводской_номер ='" + dataGridView1[3, selectedRow].Value.ToString() + "'", conn);
            conn.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Код_СИ = reader["Код_СИ"].ToString();

            }
            conn.Close();


            conn.Open();
            SqlCommand command = new SqlCommand("update Ремонт set Код_СИ = @Код_СИ, Дата_поступления = @Дата_поступления, Дата_завершения = @Дата_завершения, Неисправность = @Неисправность, Стоимость = @Стоимость  where Код_ремонта=" + changedCell, conn);


            command.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
            command.Parameters.AddWithValue("@Дата_поступления", dataGridView1[4, selectedRow].Value.ToString());
            command.Parameters.AddWithValue("@Дата_завершения", dataGridView1[5, selectedRow].Value.ToString());
            command.Parameters.AddWithValue("@Неисправность", dataGridView1[6, selectedRow].Value.ToString());
            command.Parameters.AddWithValue("@Стоимость", dataGridView1[7, selectedRow].Value.ToString());
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRow = e.RowIndex;
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            changedCell = dataGridView1[0, e.RowIndex].Value.ToString();
        }
    }
}
