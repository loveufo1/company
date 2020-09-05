using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace company
{
    public partial class change : Form
    {
        public change()
        {
            InitializeComponent();
        }
        internal string get = "";
        private void button1_Click(object sender, EventArgs e)
        {
            
            string search = comboBox2.Text;
            var n = (from m in db.TSell
                     where m.OrderNo == search
                     select m).FirstOrDefault();
            n.CDate = Convert.ToInt32(textBox9.Text);
            n.Date= Convert.ToInt32(textBox7.Text);
            n.Quantity= Convert.ToInt32(textBox2.Text);
            n.ItemNo = textBox1.Text;
            n.HarberCode = comboBox1.Text;
            n.Size = textBox3.Text;
            n.Po = textBox8.Text;
            n.NW = textBox4.Text;
            n.GW = textBox5.Text;
            n.Out = textBox6.Text;
            n.Al = Convert.ToInt32(textBox10.Text);
            db.SaveChanges();
            MessageBox.Show("修改成功");

        }
        homecompanyEntities1 db = new homecompanyEntities1();
        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void change_Load(object sender, EventArgs e)
        {
            var n1 = from m in db.TSell
                    orderby m.Date
                    
                    select m.OrderNo;
            string[] y =
            {         
                "",
                DateTime.Today.AddYears(1).Year.ToString(),
                DateTime.Today.Year.ToString(),
                DateTime.Today.AddYears(-1).Year.ToString()
            };

            
            comboBox3.DataSource = y;
            comboBox2.DataSource = n1.ToList();
            
            var n = (from m in db.TSell
                     where m.OrderNo == get
                     select m).FirstOrDefault();
            if(n != null) { 
            comboBox1.Text = n.HarberCode;
            comboBox1.SelectedItem = n.HarberCode;
            textBox1.Text = n.ItemNo;
            textBox8.Text = n.Po;
            textBox7.Text = n.Date.ToString();
            textBox9.Text = n.CDate.ToString();
            textBox2.Text = n.Quantity.ToString();
            textBox3.Text = n.Size;
            textBox4.Text = n.NW;
            textBox5.Text = n.GW;
            textBox6.Text = n.Out;
            comboBox2.SelectedItem = get;
            comboBox2.Text = get;
                textBox10.Text = n.Al.ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            string search = comboBox2.Text;
            var n = (from m in db.TSell
                     where m.OrderNo == search
                     select m).FirstOrDefault();
            db.TSell.Remove(n);
            db.SaveChanges();
            MessageBox.Show("刪除成功");

            var n2 = from m in db.TSell
                    orderby m.Date

                    select m.OrderNo;
            comboBox2.DataSource = n2.ToList();

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string search = comboBox2.Text;
            var n = (from m in db.TSell
                     where m.OrderNo == search
                     select m).FirstOrDefault();
            comboBox1.Text = n.HarberCode;
            comboBox1.SelectedItem = n.HarberCode;
            textBox1.Text = n.ItemNo;
            textBox8.Text = n.Po;
            textBox7.Text = n.Date.ToString();
            textBox9.Text = n.CDate.ToString();
            textBox2.Text = n.Quantity.ToString();
            textBox3.Text = n.Size;
            textBox4.Text = n.NW;
            textBox5.Text = n.GW;
            textBox6.Text = n.Out;
            textBox10.Text = n.Al.ToString();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            var n = from m in db.TSell
                    where m.Date.ToString().Contains(comboBox3.Text)
                    orderby m.Date
                    select m.OrderNo;

            comboBox2.DataSource = n.ToList();
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            string Filestr = $@"C:\hw4\{textBox8.Text}";

            Excel.Application app = new Excel.Application();
            Excel.Workbook exwb = app.Workbooks.Add();

            Excel.Worksheet ws = new Excel.Worksheet();
            ws = exwb.Worksheets[1];
            ws.Name = "cell";

            if (comboBox1.Text == "66")
            {
                app.Cells[3, 1] = "66";
                app.Cells[5, 1] = "LAEM CHABANG";
                app.Cells[8, 2] = textBox8.Text;
                app.Cells[19, 1] = "66";
                app.Cells[21, 1] = "LAEM CHABANG";
                app.Cells[24, 2] = textBox8.Text;

            }
            else if (comboBox1.Text == "84")
            {
                app.Cells[3, 1] = "84";
                app.Cells[5, 1] = "HO CHI MINH";
                app.Cells[8, 2] = textBox8.Text;
                app.Cells[19, 1] = "84";
                app.Cells[21, 1] = "HO CHI MINH";
                app.Cells[24, 2] = textBox8.Text;

            }
            else if (comboBox1.Text == "SMART TECH")
            {

                app.Cells[3, 1] = "SMART TECH";
                app.Cells[5, 1] = "BAVET";
                app.Cells[8, 2] = textBox8.Text;
                app.Cells[19, 1] = "SMART TECH";
                app.Cells[21, 1] = "BAVET";
                app.Cells[24, 2] = textBox8.Text;


            }




            //右上

            app.Cells[4, 1] = "       ";
            app.Cells[6, 1] = "MADE IN TAIWAN";
            app.Cells[7, 1] = "R.O.C.";
            app.Cells[8, 1] = @"PO:";
            app.Cells[9, 1] = @"C/No.:";
            app.Cells[11, 5] = textBox6.Text;
            app.Cells[11, 1] = "A";

            app.Cells[4, 5] = "ARTCLE:STEEL PIN";
            app.Cells[5, 5] = "ITEM NO. ";
            app.Cells[5, 6] = textBox1.Text;
            app.Cells[6, 5] = "SIZE ";
            app.Cells[6, 6] = textBox2.Text;
            app.Cells[7, 5] = "QNTY ";
            app.Cells[7, 6] = textBox3.Text + "  PC";
            app.Cells[8, 5] = "N.W";
            app.Cells[8, 6] = textBox4.Text + "  KG";
            app.Cells[9, 5] = "G.W";
            app.Cells[9, 6] = textBox5.Text + "  KG";
            app.Cells[10, 5] = textBox7.Text;
            app.Cells[11, 5] = textBox10.Text;
            //右下

            app.Cells[20, 1] = "       ";
            app.Cells[22, 1] = "MADE IN TAIWAN";
            app.Cells[23, 1] = "R.O.C.";
            app.Cells[24, 1] = @"PO:";
            app.Cells[25, 1] = @"C/No.:";
            app.Cells[27, 5] = textBox6.Text;
            app.Cells[27, 1] = "B";

            app.Cells[20, 5] = "ARTCLE:STEEL PIN";
            app.Cells[21, 5] = "ITEM NO. ";
            app.Cells[21, 6] = textBox1.Text;
            app.Cells[22, 5] = "SIZE ";
            app.Cells[22, 6] = textBox2.Text;
            app.Cells[23, 5] = "QNTY ";
            app.Cells[23, 6] = textBox3.Text + "  PC";
            app.Cells[24, 5] = "N.W";
            app.Cells[24, 6] = textBox4.Text + "  KG";
            app.Cells[25, 5] = "G.W";
            app.Cells[25, 6] = textBox5.Text + "  KG";
            app.Cells[26, 5] = textBox7.Text;
            app.Cells[27, 5] = textBox10.Text;


            exwb.SaveAs(Filestr);

            ws = null;
            exwb.Close();
            exwb = null;
            app.Quit();
            app = null;
        }

        

        private void comboBox2_MouseLeave(object sender, EventArgs e)
        {
            string search = comboBox2.Text;
            var n = (from m in db.TSell
                     where m.OrderNo == search
                     select m).FirstOrDefault();
            comboBox1.Text = n.HarberCode;
            comboBox1.SelectedItem = n.HarberCode;
            textBox1.Text = n.ItemNo;
            textBox8.Text = n.Po;
            textBox7.Text = n.Date.ToString();
            textBox9.Text = n.CDate.ToString();
            textBox2.Text = n.Quantity.ToString();
            textBox3.Text = n.Size;
            textBox4.Text = n.NW;
            textBox5.Text = n.GW;
            textBox6.Text = n.Out;
            textBox10.Text = n.Al.ToString();
        }
    }
}
