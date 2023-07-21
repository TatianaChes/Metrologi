using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class План : MaterialForm
    {
        string connection = @"Data Source = .\SQLEXPRESS; Initial Catalog = Метрология; Integrated Security = True";
        int selectedRow;
        string changedCell;
        
        string Код_СИ, Код_вида;
        public План()
        {
            InitializeComponent();
        }

        private void План_Load(object sender, EventArgs e)
        {
            Print();
        }

        private void Print()
        {
            SqlConnection conn = new SqlConnection(connection);
            SqlDataAdapter adapter = new SqlDataAdapter(" select Код_плана as Код, Наименование_СИ.Наименование, Тип_инструмента.Тип, Средство_измерения.Заводской_номер,Вид_метрологического_контроля.Наименование as Контроль, МПИ, Дата_следующей_поверки as Следующая_поверка from План_контроля join Средство_измерения on Средство_измерения.Код_СИ = План_контроля.Код_СИ join Наименование_СИ on Наименование_СИ.Код_наименования = Средство_измерения.Код_наименования left join Тип_инструмента on Тип_инструмента.Код_типа = Средство_измерения.Код_типа join Вид_метрологического_контроля on Вид_метрологического_контроля.Код_вида = План_контроля.Код_вида", conn);
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
            dataGridView1.Columns[6].Width = 120;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Добавить_в_план в_план = new Добавить_в_план();
            в_план.Show(); 
            в_план.FormClosed += new FormClosedEventHandler(af_FormClosed);
        }

        private void af_FormClosed(object sender, FormClosedEventArgs e)
        {
            Print();
        }

        private void button3_Click(object sender, EventArgs e) // удаление
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
                    cmd = new SqlCommand("DELETE FROM План_контроля WHERE Код_плана = @Код_плана", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@Код_плана", dataGridView1[0, selectedRow].Value.ToString());
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

        private void button2_Click(object sender, EventArgs e) // редактирование
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

            SqlCommand cmd2 = new SqlCommand("select Код_вида from Вид_метрологического_контроля where  Наименование='" + dataGridView1[4, selectedRow].Value.ToString() + "'", conn);
            conn.Open();
            SqlDataReader reader2 = cmd2.ExecuteReader();
            if (reader2.Read())
            {
                Код_вида = reader2["Код_вида"].ToString();

            }
            conn.Close();

            conn.Open();
            SqlCommand command = new SqlCommand("update План_контроля set Код_СИ = @Код_СИ, Код_вида = @Код_вида, МПИ = @МПИ, Дата_следующей_поверки = @Дата_следующей_поверки  where Код_плана=" + changedCell, conn);


            command.Parameters.AddWithValue("@Код_СИ", SqlDbType.Int).Value = Convert.ToInt32(Код_СИ);
            command.Parameters.AddWithValue("@Код_вида", SqlDbType.Int).Value = Convert.ToInt32(Код_вида);
            command.Parameters.AddWithValue("@МПИ", dataGridView1[5, selectedRow].Value.ToString());
            command.Parameters.AddWithValue("@Дата_следующей_поверки", dataGridView1[6, selectedRow].Value.ToString());
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

            Microsoft.Office.Interop.Excel.Range xlSheetRange;
            //Приложение
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //выбираем лист на котором будем работать (Лист 1)
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.Sheets[1];
            //Название листа
            ExcelWorkSheet.Name = "Данные";
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                       

            //делаем полужирный текст и перенос слов
            xlSheetRange = ExcelWorkSheet.get_Range("A1:Z1", Type.Missing);
            xlSheetRange.Font.Bold = true;
            // Выделяем диапазон ячеек 
            Microsoft.Office.Interop.Excel.Range _excelCells2 = (Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet.get_Range("A1", "Z1").Cells;
            // Производим объединение
            _excelCells2.Merge(Type.Missing);
            xlSheetRange.Cells[1, 1] = "План метрологического контроля";
            (xlSheetRange.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range).Font.Size = 14;

            // Шапка таблицы
            ExcelApp.Cells[4, 1] = "№";
            ExcelApp.Cells[4, 2] = "Наименование";
            ExcelApp.Cells[4, 3] = "Тип";
            ExcelApp.Cells[4, 4] = "Заводской номер";
            ExcelApp.Cells[4, 5] = "Контроль";
            ExcelApp.Cells[4, 6] = "МПИ";
            ExcelApp.Cells[4, 7] = "Следующая поверка";
            xlSheetRange.Cells[4, 1].Font.Bold = true;
            xlSheetRange.Cells[4, 2].Font.Bold = true;
            xlSheetRange.Cells[4, 3].Font.Bold = true;
            xlSheetRange.Cells[4, 4].Font.Bold = true;
            xlSheetRange.Cells[4, 5].Font.Bold = true;
            xlSheetRange.Cells[4, 6].Font.Bold = true;
            xlSheetRange.Cells[4, 7].Font.Bold = true;
            ExcelWorkSheet.get_Range("A4:G4", Type.Missing).Interior.Color = Color.Green;
            ExcelWorkSheet.get_Range("A4:G4", Type.Missing).Font.Color = Color.White;

            // Перенос сведений из датагрида
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 5, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }

            ExcelWorkSheet.get_Range("A4", "G21").Cells.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            
            xlSheetRange = ExcelWorkSheet.UsedRange;

            //выравниваем строки и колонки по их содержимому
            xlSheetRange.Columns.AutoFit();
            xlSheetRange.Rows.AutoFit();

            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;


            ExcelWorkBook.SaveAs(@"C:\sql\exel.xlsx");  // сохранение
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            changedCell = dataGridView1[0, e.RowIndex].Value.ToString();
           
        }
    }
}
