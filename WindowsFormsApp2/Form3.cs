using System;
using System.Globalization;
using MaterialSkin.Controls;

namespace WindowsFormsApp2
{
    public partial class Form3 : MaterialForm
    {
        int month, year;

        public static int static_month, static_year;
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            displayDays();
            UserControlDays user = new UserControlDays();
            user.displayEvent();
            user.displayEvent2();
        }

        private void displayDays()
        {

            DateTime now = DateTime.Now;
            month = now.Month;
            year = now.Year;

            String monthname = DateTimeFormatInfo.CurrentInfo.GetMonthName(month);
            label8.Text = monthname + " " + year;


            static_month = month;
            static_year = year;

            // first day of the month
            DateTime startofmonth = new DateTime(year, month, 1);
            // count of days of the month
            int days = DateTime.DaysInMonth(year, month);
            int dayoftheweek = Convert.ToInt32(startofmonth.DayOfWeek.ToString("d")) + 1;

            for (int i = 1; i < dayoftheweek; i++) {
                UserControlBlank ucblack = new UserControlBlank();
                flowLayoutPanel1.Controls.Add(ucblack);
            }
            for (int i = 1; i <= days; i++)
            {
                UserControlDays ucdays = new UserControlDays();
                ucdays.days(i);
                flowLayoutPanel1.Controls.Add(ucdays);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            if (month != 1)
            {
                month--;
            }
            else
            {
                month = 12;
                year--;
            }

            static_month = month;
            static_year = year;

            String monthname = DateTimeFormatInfo.CurrentInfo.GetMonthName(month);
            label8.Text = monthname + " " + year;

            DateTime now = DateTime.Now;
            // first day of the month
            DateTime startofmonth = new DateTime(year, month, 1);
            // count of days of the month
            int days = DateTime.DaysInMonth(year, month);
            int dayoftheweek = Convert.ToInt32(startofmonth.DayOfWeek.ToString("d")) + 1;

            for (int i = 1; i < dayoftheweek; i++)
            {
                UserControlBlank ucblack = new UserControlBlank();
                flowLayoutPanel1.Controls.Add(ucblack);
            }
            for (int i = 1; i <= days; i++)
            {
                UserControlDays ucdays = new UserControlDays();
                ucdays.days(i);
                flowLayoutPanel1.Controls.Add(ucdays);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();


            if (month != 12)
            {
                month++;
            }
            else
            {
                month = 1;
                year++;
            }

            static_month = month;
            static_year = year;

            String monthname = DateTimeFormatInfo.CurrentInfo.GetMonthName(month);
            label8.Text = monthname + " " + year;

            DateTime now = DateTime.Now;
            // first day of the month
            DateTime startofmonth = new DateTime(year, month, 1);
            // count of days of the month
            int days = DateTime.DaysInMonth(year, month);
            int dayoftheweek = Convert.ToInt32(startofmonth.DayOfWeek.ToString("d")) + 1;

            for (int i = 1; i < dayoftheweek; i++)
            {
                UserControlBlank ucblack = new UserControlBlank();
                flowLayoutPanel1.Controls.Add(ucblack);
            }
            for (int i = 1; i <= days; i++)
            {
                UserControlDays ucdays = new UserControlDays();
                ucdays.days(i);
                flowLayoutPanel1.Controls.Add(ucdays);
            }
        }
    }
}
