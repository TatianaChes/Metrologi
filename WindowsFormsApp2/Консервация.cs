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
    public partial class Консервация : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        int selectedRow;
        string Код_СИ, changedCell, Фамилия, Имя, Отчество, Должность;
        int i = 0;
        public Консервация()
        {
            InitializeComponent();
        }

        private void Консервация_Load(object sender, EventArgs e)
        {
            Print();

        }

        private void Print()
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter("select Код_консервации as Код, Наименование_СИ.Наименование, Тип_инструмента.Тип, Средство_измерения.Заводской_номер," +
                " Дата_начала_хранения as Начало_хранения, Дата_окончания_хранения as Окончание, Условия_хранения as Условия from Инструмент_в_консервации join Средство_измерения on" +
                " Средство_измерения.Код_СИ = Инструмент_в_консервации.Код_СИ join Наименование_СИ on Наименование_СИ.Код_наименования " +
                "= Средство_измерения.Код_наименования  left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа", conn);
            SqlCommandBuilder builder = new SqlCommandBuilder(adapter);
            System.Data.DataTable table = new System.Data.DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;

            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[5].Width = 70;
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
                    cmd = new SqlCommand("DELETE FROM Инструмент_в_консервации WHERE Код_консервации = @Код_консервации", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@Код_консервации", dataGridView1[0, selectedRow].Value.ToString());
                    cmd.ExecuteNonQuery();
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

        private void button1_Click(object sender, EventArgs e)
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

            SqlCommand cmd2 = new SqlCommand("select count(*) from Инструмент_в_консервации", conn);
            conn.Open();
            int count = (int)cmd2.ExecuteScalar();
            conn.Close();

            SqlCommand cmd3 = new SqlCommand("insert into Инструмент_в_консервации(Код_консервации, Код_СИ, Дата_начала_хранения, Дата_окончания_хранения, Условия_хранения) values (@Код_консервации, @Код_СИ, @Дата_начала_хранения, @Дата_окончания_хранения, @Условия_хранения)", conn);

            conn.Open();
            cmd3.Parameters.AddWithValue("@Код_консервации", SqlDbType.Int).Value = count + 1;
            cmd3.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
            cmd3.Parameters.AddWithValue("@Дата_начала_хранения", dataGridView1[4, selectedRow].Value.ToString());
            cmd3.Parameters.AddWithValue("@Дата_окончания_хранения", dataGridView1[5, selectedRow].Value.ToString());
            cmd3.Parameters.AddWithValue("@Условия_хранения", dataGridView1[6, selectedRow].Value.ToString());
            try
            {
                cmd3.ExecuteNonQuery();
                MessageBox.Show("Инструмент в консервации!");
            }
            catch
            { MessageBox.Show("Не удалось законсервировать инструмент!"); }
            finally { conn.Close(); Print(); }

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
            SqlCommand command = new SqlCommand("update Инструмент_в_консервации set Код_СИ = @Код_СИ, Дата_начала_хранения = @Дата_начала_хранения, Дата_окончания_хранения = @Дата_окончания_хранения, Условия_хранения = @Условия_хранения  where Код_консервации=" + changedCell, conn);


            command.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
            command.Parameters.AddWithValue("@Дата_начала_хранения", dataGridView1[4, selectedRow].Value.ToString());
            command.Parameters.AddWithValue("@Дата_окончания_хранения", dataGridView1[5, selectedRow].Value.ToString());
            command.Parameters.AddWithValue("@Условия_хранения", dataGridView1[6, selectedRow].Value.ToString());
           
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

        private void button4_Click(object sender, EventArgs e)
        {

            SqlConnection conn = new SqlConnection(connection);
            conn.Open();
            SqlCommand command = new SqlCommand("select Фамилия, left(Имя,1) + '.' as Имя, left(Отчество,1) + '.' as Отчество, Должность.Наименование_должности from Комиссия join Должность on Должность.Код_должности = Комиссия.Код_должности where Наименование_должности = 'Инженер по метрологии'",conn);
            SqlDataReader reader = command.ExecuteReader();

            if (reader.Read()) {
                Фамилия = reader["Фамилия"].ToString();
                Имя = reader["Имя"].ToString();
                Отчество = reader["Отчество"].ToString();
                Должность = reader["Наименование_должности"].ToString();

            }
            conn.Close();

            try
            {
                Microsoft.Office.Interop.Word.Application winword =
                    new Microsoft.Office.Interop.Word.Application();

                winword.Visible = false;

                //Заголовок документа
                winword.Documents.Application.Caption = "Акт консервации";

                object missing = System.Reflection.Missing.Value;

                //Создание нового документа
                Microsoft.Office.Interop.Word.Document document =
                    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                //Добавление верхнего колонтитула
                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[
                    Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add( headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.Size = 8;
                    headerRange.Text = 
                        "ООО 'Инженерные Технологии' 123312 г. Москва ул. Красногвардейская 3-я, д.3, офис 70." + Environment.NewLine
                        + "Тел. +7(495)145-8888, +7(3532)666-777, +7(3532)666-888, e-mail: Info@e-t-a.org, www-e-t-a.org" + Environment.NewLine
                        + "ИНН/КПП 560908880/770301001; ОГРН 1125658026680" + Environment.NewLine
                        + "р/с 40702810900330000314, к/с 30101810000000000805" + Environment.NewLine
                        + "ПАО 'АК БАРС' БАНК, г. Казань, БИК 0 49205805";
                }

                //Добавление текста в документ
    
                //Добавление текста со стилем Заголовок 1
                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                para1.Range.Font.Size = 14;
                para1.Range.Text = Environment.NewLine + Environment.NewLine + "Акт консервации №____  от " + DateTime.Now.ToString("dd.MM.yyyy");
                para1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para1.Range.InsertParagraphAfter();

                int count = dataGridView1.DisplayedRowCount(true);
                //Создание таблицы 5х6
                Table firstTable = document.Tables.Add(para1.Range, count, 6, ref missing, ref missing);

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

                para1.Range.Font.Size = 14;
                para1.Range.Text = Environment.NewLine + Environment.NewLine + "Консервацию провел" + " "+ Должность +  " " + Фамилия +  " "+ Имя + " "+ Отчество +  "   _________";
                para1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                winword.Visible = true;

                document.SaveAs(@"C:\sql\1.docx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

      
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            changedCell = dataGridView1[0, e.RowIndex].Value.ToString();
        }
    }
}
