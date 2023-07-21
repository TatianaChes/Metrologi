using MaterialSkin.Controls;
using Microsoft.Office.Interop.Word;
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
    public partial class Списанный : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        string Код_СИ, Фамилия, Имя, Отчество, Должность, Фамилия2, Имя2, Отчество2, Должность2, Фамилия3, Имя3, Отчество3, Должность3;
        int selectedRow;
        public Списанный()
        {
            InitializeComponent();
        }

        private void Списанный_Load(object sender, EventArgs e)
        {
            Print();

        }

        private void Print()
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter(" select Код_списания as Код, Наименование_СИ.Наименование, Тип_инструмента.Тип, Средство_измерения.Заводской_номер, Причина_списания, Дата_списания from Списанный_инструмент join Средство_измерения on Средство_измерения.Код_СИ = Списанный_инструмент.Код_СИ join Наименование_СИ on Наименование_СИ.Код_наименования = Средство_измерения.Код_наименования  left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа", conn);
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            System.Data.DataTable table = new System.Data.DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;

            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[3].Width = 110;
            dataGridView1.Columns[4].Width = 110;
            dataGridView1.Columns[5].Width = 90;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRow = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[selectedRow];

            textBox3.Text = row.Cells[3].Value.ToString();
            textBox4.Text = row.Cells[4].Value.ToString();
            textBox5.Text = row.Cells[5].Value.ToString();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlCommand cmd = new SqlCommand("select Код_СИ from Средство_измерения where Заводской_номер ='" + textBox3.Text + "'", conn);
            conn.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Код_СИ = reader["Код_СИ"].ToString();

            }
            conn.Close();

            SqlCommand cmd2 = new SqlCommand("select count(*) from Списанный_инструмент", conn);
            conn.Open();
            int count = (int)cmd2.ExecuteScalar();
            conn.Close();

            SqlCommand cmd3 = new SqlCommand("insert into Списанный_инструмент(Код_списания, Код_СИ, Причина_списания, Дата_списания," +
                "Код_начальника_производства, Код_начальника_цеха, Код_инженера_по_метрологии) values (@Код_списания, @Код_СИ, @Причина_списания, @Дата_списания,@Код_начальника_производства, @Код_начальника_цеха, @Код_инженера_по_метрологии)", conn);

            conn.Open();
            cmd3.Parameters.AddWithValue("@Код_списания", SqlDbType.Int).Value = count + 1;
            cmd3.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
            cmd3.Parameters.AddWithValue("@Причина_списания", textBox4.Text);
            cmd3.Parameters.AddWithValue("@Дата_списания", textBox5.Text);
            cmd3.Parameters.AddWithValue("@Код_начальника_производства", SqlDbType.Int).Value = 1;
            cmd3.Parameters.AddWithValue("@Код_начальника_цеха", SqlDbType.Int).Value = 2;
            cmd3.Parameters.AddWithValue("@Код_инженера_по_метрологии", SqlDbType.Int).Value = 3;



            try
            {
                cmd3.ExecuteNonQuery();
                MessageBox.Show("Инструмент списан!");
            }
            catch
            { MessageBox.Show("Не удалось списать инструмент!"); }
            finally { conn.Close(); Print(); }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlCommand cmd;
            SqlConnection con;
           
            con = new SqlConnection(connection);
            con.Open();
            cmd = new SqlCommand(" Update Средство_измерения set Заводской_номер = @Заводской_номер  where Средство_измерения.Код_СИ = ( select Код_СИ from Списанный_инструмент  where Код_списания = " + dataGridView1[0, selectedRow].Value.ToString() + ")", con);

            cmd.Parameters.AddWithValue("@Заводской_номер", textBox3.Text);

            try
            {
                cmd.ExecuteNonQuery();
                MessageBox.Show("Данные обновлены!");
            }
            catch
            { MessageBox.Show("Не удалось обновить данные!"); }
           

           
            string quet = "Update Списанный_инструмент set Причина_списания = @Причина_списания, Дата_списания = @Дата_списания  Where Код_списания = " + dataGridView1[0, selectedRow].Value.ToString();
            cmd = new SqlCommand(quet, con);
         
            cmd.Parameters.AddWithValue("@Причина_списания", textBox4.Text);
            cmd.Parameters.AddWithValue("@Дата_списания", textBox5.Text);

            try
            {
                cmd.ExecuteNonQuery();
                MessageBox.Show("Данные обновлены!");
            }
            catch
            { MessageBox.Show("Не удалось обновить данные!"); }
            finally { con.Close(); Print(); }

           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(connection);
            conn.Open();
            SqlCommand command = new SqlCommand("select Фамилия, left(Имя,1) + '.' as Имя, left(Отчество,1) + '.' as Отчество, Должность.Наименование_должности from Комиссия join Должность on Должность.Код_должности = Комиссия.Код_должности where Наименование_должности = 'Начальник производства'", conn);
            SqlDataReader reader = command.ExecuteReader();

            if (reader.Read())
            {
                Фамилия = reader["Фамилия"].ToString();
                Имя = reader["Имя"].ToString();
                Отчество = reader["Отчество"].ToString();
                Должность = reader["Наименование_должности"].ToString();

            }
            conn.Close();

            SqlCommand command2 = new SqlCommand("select Фамилия, left(Имя,1) + '.' as Имя, left(Отчество,1) + '.' as Отчество, Должность.Наименование_должности from Комиссия join Должность on Должность.Код_должности = Комиссия.Код_должности where Наименование_должности = 'Начальник цеха'", conn);
            conn.Open();
            SqlDataReader reader2 = command2.ExecuteReader();

            if (reader2.Read())
            {
                Фамилия2 = reader2["Фамилия"].ToString();
                Имя2 = reader2["Имя"].ToString();
                Отчество2 = reader2["Отчество"].ToString();
                Должность2 = reader2["Наименование_должности"].ToString();

            }
            conn.Close();

            
            SqlCommand command3 = new SqlCommand("select Фамилия, left(Имя,1) + '.' as Имя, left(Отчество,1) + '.' as Отчество, Должность.Наименование_должности from Комиссия join Должность on Должность.Код_должности = Комиссия.Код_должности where Наименование_должности = 'Инженер по метрологии'", conn);
            conn.Open();
            SqlDataReader reader3 = command3.ExecuteReader();

            if (reader3.Read())
            {
                Фамилия3 = reader3["Фамилия"].ToString();
                Имя3 = reader3["Имя"].ToString();
                Отчество3 = reader3["Отчество"].ToString();
                Должность3 = reader3["Наименование_должности"].ToString();

            }
            conn.Close();

            SqlCommand cmd2 = new SqlCommand("select count(*) from Списанный_инструмент", conn);
            conn.Open();
            int count2 = (int)cmd2.ExecuteScalar();
            conn.Close();

            try
            {
                Microsoft.Office.Interop.Word.Application winword =
                    new Microsoft.Office.Interop.Word.Application();

                winword.Visible = false;

                //Заголовок документа
                winword.Documents.Application.Caption = "Акт списания";

                object missing = System.Reflection.Missing.Value;

                //Создание нового документа
                Microsoft.Office.Interop.Word.Document document =
                    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

             
                //Добавление текста в документ

                //Добавление текста со стилем Заголовок 1
                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                para1.Range.Font.Size = 14;
                para1.Range.Font.Bold = 1;
                para1.Range.Text = "\t\t\t\t\t\t\t\tАкт №  ____  от " + DateTime.Now.ToString("dd") + " " + DateTime.Now.ToString("MMMM") + " "+ DateTime.Now.ToString("yyyy") + " г." + Environment.NewLine + "\t\t\t\t\tна списание малоценных и быстроизнашивающихся предметов" + Environment.NewLine + Environment.NewLine + "Организация: ООО 'Инженерные Технологии'";
                para1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para1.Range.InsertParagraphAfter();

                int count = dataGridView1.DisplayedRowCount(true);
                //Создание таблицы 5х6
                Table firstTable = document.Tables.Add(para1.Range, count, 5, ref missing, ref missing);

                firstTable.Borders.Enable = 1;

                foreach (Row row in firstTable.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        //Заголовок таблицы
                        if (cell.RowIndex == 1)
                        {

                            cell.Range.Text = dataGridView1.Columns[cell.ColumnIndex].HeaderText;
                            cell.Range.Font.Bold = 1;
                            //Задаем шрифт и размер текста
                            cell.Range.Font.Name = "verdana";
                            cell.Range.Font.Size = 10;
                            cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                            //Выравнивание текста в заголовках столбцов по центру
                            cell.VerticalAlignment =
                            WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment =
                            WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        //Значения ячеек
                        else
                        {

                            cell.Range.Text = dataGridView1.Rows[cell.RowIndex - 2].Cells[cell.ColumnIndex].Value.ToString();
                        }

                    }

                }
              
                Microsoft.Office.Interop.Word.Paragraph para3 = document.Content.Paragraphs.Add(ref missing);
                para3.Range.Font.Size = 14;
                para3.Range.Text = Environment.NewLine + "Члены комиссии: " + "\t" + Должность + " " + Фамилия + " " + Имя + " " + Отчество + " " + "________" + "\t\t" + "________________________ " + Environment.NewLine + "\t\t\t\t" + Должность2 + " " + Фамилия2 + " " + Имя2 + " " + Отчество2 + "\t\t" + "________" + "\t\t\t" + "________________________ " + Environment.NewLine + "\t\t\t\t\t" + Должность3 + " " + Фамилия3 + " " + Имя3 + " " + Отчество3 + " " + "________" + "\t\t" + "________________________ ";
                para3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                winword.Visible = true;

                document.SaveAs(@"C:\sql\2.docx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
