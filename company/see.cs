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
            string year = yearC.Text;
            var n = from m in db.TSell
                    where m.Date.ToString().Contains(year)
                    select new
                    {
                        m.Size,
                        m.Quantity
                    };
            var n1 = n.GroupBy(m => m.Size, m => m.Quantity, (Size, quantity) => new
            {
                規格 = Size,
                數量 = quantity.Sum()
            });
            dataGridView1.DataSource = n1.ToList();
        }

        private void see_Load(object sender, EventArgs e)
        {
            string[] y =
            {         "",
                DateTime.Today.AddYears(1).Year.ToString(),
                DateTime.Today.Year.ToString(),
                DateTime.Today.AddYears(-1).Year.ToString()
            };
            yearC.DataSource = y;
            yearC.SelectedItem = DateTime.Today.Year.ToString();

            monthC.SelectedItem = DateTime.Now.Month.ToString();
            dataGridView1.Font = new Font("標楷體", 17);
        }
        homecompanyEntities1 db = new homecompanyEntities1();
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string date = yearC.Text+monthC.Text;

            var n = from m in db.TSell
                    where m.Date.ToString().Contains(date)
                    select new
                    {
                        m.Size,
                        m.Quantity
                    };
            var n1 = n.GroupBy(m => m.Size, m => m.Quantity, (Size, quantity) => new
            {
                規格 = Size,
                數量 = quantity.Sum()
            });
            dataGridView1.DataSource = n1.ToList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            monthC.Text = " ";
            var n = from m in db.TSell
                    where m.Date.ToString().Contains(yearC.Text)
                    select new
                    {
                        m.Size,
                        m.Quantity
                    };
            var n1 = n.GroupBy(m => m.Size, m => m.Quantity, (Size, quantity) => new
            {
                規格 = Size,
                數量 = quantity.Sum()
            });
            dataGridView1.DataSource = n1.ToList();
            
        }
    }
}
