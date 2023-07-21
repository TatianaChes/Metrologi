using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace WindowsFormsApp2
{
    public partial class Form1 :  MaterialForm
    {
        private string connectionString;
        private OleDbConnection connect;
        static int count = 0;
        public Form1()
        {
            InitializeComponent();


            // установка стиля

            MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;

            // установка цета

            materialSkinManager.ColorScheme = new ColorScheme(Primary.Grey800, Primary.Grey800, Primary.Grey900, Accent.Purple700, TextShade.WHITE);
            SetStyle(ControlStyles.SupportsTransparentBackColor | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true) {
                connectionString = @"Provider=SQLOLEDB; DRIVER={SQL Server}; Data Source=.\SQLEXPRESS; Initial Catalog=Метрология; Integrated Security=SSPI;";
                OpenConnect();
            }
            else {
                OpenConnect();
            }
            
        }

        private void OpenConnect()
        {
            connect = new OleDbConnection(connectionString);
            try
            {
               
                connect.Open(); // открытие БД
                this.Hide();
                Form2 af = new Form2();
                af.Owner = this;
                af.Show(); 
                af.button5.Visible = false;

            }
            catch
            {
                try
                {

                    SqlConnection con = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=Метрология;Integrated Security=True");
                    SqlDataAdapter adapter = new SqlDataAdapter("Select count(*) from Оператор  where Логин = '" + textBox1.Text + "' and Пароль = '" +
                       GetHash(textBox2.Text) + "'", con);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    if (dt.Rows[0][0].ToString() == "1")
                    {
                        this.Hide();
                        Form2 af = new Form2();
                        af.Owner = this;
                        af.Show();
                        af.button11.Visible = false;
                        af.button12.Visible = false;

                    }

                    else
                    {
                        MessageBox.Show("Неверный логин или пароль", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        count++;

                        if (count == 3)
                        {
                            count = 0;
                            Form4 af = new Form4();
                            af.Show();
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Неверный логин или пароль", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    count++;

                    if (count == 3)
                    {
                        count = 0;
                        Form4 af = new Form4();
                        af.Show();
                    }
                }
            }
        }

        public string GetHash(string input)
        {
            var md5 = MD5.Create();
            var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));

            return Convert.ToBase64String(hash);
        }
    }
}
