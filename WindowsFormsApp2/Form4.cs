using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
            Load += (sender, args) => StartTimer();
        }

        private async void StartTimer()
        {
            TimeSpan ts = new TimeSpan(0, 1, 0);
            while (ts > TimeSpan.Zero)
            {
                label1.Text = ts.ToString();
                await Task.Delay(1000);
                ts -= TimeSpan.FromSeconds(1);
            }
            Close();
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }
    }
}
