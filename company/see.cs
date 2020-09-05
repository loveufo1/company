using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace company
{
    public partial class see : Form
    {
        public see()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void see_Load(object sender, EventArgs e)
        {
            string[] y =
            {         "",
                DateTime.Today.AddYears(1).Year.ToString(),
                DateTime.Today.Year.ToString(),
                DateTime.Today.AddYears(-1).Year.ToString()
            };
            comboBox1.DataSource = y;
            comboBox1.SelectedItem = DateTime.Today.Year.ToString();

            comboBox2.SelectedItem = DateTime.Now.Month.ToString();
        }
    }
}
