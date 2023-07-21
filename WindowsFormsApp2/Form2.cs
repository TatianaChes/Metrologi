using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace WindowsFormsApp2
{
    public partial class Form2 : MaterialForm
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e) // календарь
        {

            Form3 form3 = new Form3();
            form3.Show();

        }

        private void button8_Click(object sender, EventArgs e)// на главную
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Close();
        }

        private void button9_Click(object sender, EventArgs e)// выход
        {
            var result = MessageBox.Show("Вы действительно хотите завершить работу?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
            {

            }
            else
            {
                this.Close();
                System.Windows.Forms.Application.Exit();
            }
        }

        private void button2_Click(object sender, EventArgs e)// производство
        {
            Производство производство = new Производство();
            производство.Show();
        }

        private void button6_Click(object sender, EventArgs e)//списанные инструменты
        {
            Списанный списанный = new Списанный();
            списанный.Show();
        }

        private void button4_Click(object sender, EventArgs e)// в ремонте
        {
            В_ремонте ремонте = new В_ремонте();
            ремонте.Show();
        }

        private void button7_Click(object sender, EventArgs e)// ремонт
        {
            Ремонт ремонт = new Ремонт();
            ремонт.Show();
        }

        public void button5_Click(object sender, EventArgs e)// план контроля
        {
            План план = new План();
            план.Show();

        }

        private void button3_Click(object sender, EventArgs e)// консервация
        {
            Консервация консервация = new Консервация();
            консервация.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            СИ си = new СИ();
            си.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            План план = new План();
            план.Show();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Операторы операторы = new Операторы();
            операторы.Show();
        }
    }
}
