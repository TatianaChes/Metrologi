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
    public partial class В_ремонте : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        int selectedRow;
        string Код_СИ, changedCell;
        public В_ремонте()
        {
            InitializeComponent();
        }

        private void В_ремонте_Load(object sender, EventArgs e)
        {

            Print();
        }

        private void Print()
        {

            string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";

            SqlConnection conn = new SqlConnection(connection);

            SqlDataAdapter adapter = new SqlDataAdapter("select Код_инструмента_в_ремонте as Код, Наименование_СИ.Наименование, Тип_инструмента.Тип, Средство_измерения.Заводской_номер," +
                "Дата_обнаружения_неисправности as Дата_обнаружения, Неисправность from Инструмент_в_ремонте join Средство_измерения on Средство_измерения.Код_СИ" +
                " = Инструмент_в_ремонте.Код_СИ join Наименование_СИ on Наименование_СИ.Код_наименования = " +
                "Средство_измерения.Код_наименования  left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа", conn);
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;


            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[3].Width = 115;
            dataGridView1.Columns[4].Width = 115;
            dataGridView1.Columns[5].Width = 144;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Добавить_в_ремонт в_Ремонт = new Добавить_в_ремонт();
            в_Ремонт.Show();
            в_Ремонт.FormClosed += new FormClosedEventHandler(af_FormClosed);
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
                    cmd = new SqlCommand("DELETE FROM Инструмент_в_ремонте WHERE Код_инструмента_в_ремонте = @Код_инструмента_в_ремонте", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@Код_инструмента_в_ремонте", dataGridView1[0, selectedRow].Value.ToString());
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Удаление прошло успешно");
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
            SqlCommand command = new SqlCommand("update Инструмент_в_ремонте set Код_СИ = @Код_СИ, Дата_обнаружения_неисправности = @Дата_обнаружения_неисправности, Неисправность = @Неисправность  where Код_инструмента_в_ремонте=" + changedCell, conn);


            command.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
           command.Parameters.AddWithValue("@Дата_обнаружения_неисправности", dataGridView1[4, selectedRow].Value.ToString());
            command.Parameters.AddWithValue("@Неисправность", dataGridView1[5, selectedRow].Value.ToString());
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
        }
    }
}
