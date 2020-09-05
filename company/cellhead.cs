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
    public partial class cellhead : Form
    {
        public cellhead()
        {
            InitializeComponent();
        }
        homecompanyEntities1 db = new homecompanyEntities1();
        private void button1_Click(object sender, EventArgs e)
        {
            //處理港口
            string har="";
            if (comboBox1.Text == "66")
            {
                har = "LAEM CHABANG";
            }
            else if(comboBox1.Text == "84")
            {
                har = "HO CHI MINH";
            }
            else if (comboBox1.Text == "SMART TECH")
            {
                har = "BAVET";
            }

            //處理PO
            
            if(textBox5.Text == "")
            {
                textBox5.Text = "0";
            }

            if (textBox10.Text == "")
            {
                textBox10.Text = "0";
            }
            var neww = new TSell
            {
                HarberCode = comboBox1.Text,
                ItemNo = textBox1.Text,
                OrderNo = textBox6.Text,
                Date = Convert.ToInt32(textBox7.Text),
                Size = textBox2.Text,
                Quantity = Convert.ToInt32(textBox3.Text),
                NW = textBox4.Text,
                GW = textBox5.Text,
                Harber = har,
                Po = textBox8.Text,
                CDate = Convert.ToInt32(textBox9.Text),
                Out = "未出貨",
                Al = Convert.ToInt32(textBox10.Text)
            };
            db.TSell.Add(neww);
            db.SaveChanges();
            MessageBox.Show("新增成功");

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
                app.Cells[8, 2] =textBox8.Text;
                app.Cells[19, 1] = "66";
                app.Cells[21, 1] = "LAEM CHABANG";
                app.Cells[24, 2] = textBox8.Text;
              
            }
            else if (comboBox1.Text == "84")
            {
                app.Cells[3, 1] = "84";
                app.Cells[5, 1] = "HO CHI MINH";
                app.Cells[8, 2] =  textBox8.Text;
                app.Cells[19, 1] = "84";
                app.Cells[21, 1] = "HO CHI MINH";
                app.Cells[24, 2] = textBox8.Text;
              
            }
            else if (comboBox1.Text == "SMART TECH")
            {
                
                app.Cells[3, 1] = "SMART TECH";
                app.Cells[5, 1] = "BAVET";
                app.Cells[8, 2] =  textBox8.Text;
                app.Cells[19, 1] = "SMART TECH";
                app.Cells[21, 1] = "BAVET";
                app.Cells[24, 2] =  textBox8.Text;
                
              
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
            app.Cells[5, 6] =textBox1.Text;
            app.Cells[6, 5] = "SIZE ";
            app.Cells[6, 6] =textBox2.Text;
            app.Cells[7, 5] = "QNTY " ;
            app.Cells[7, 6] =textBox3.Text +"  PC";
            app.Cells[8, 5] = "N.W";
            app.Cells[8, 6] =  textBox4.Text+"  KG";
            app.Cells[9, 5] = "G.W";
            app.Cells[9, 6] = textBox5.Text+"  KG";
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

        private void cellhead_Load(object sender, EventArgs e)
        {
            textBox9.Text = DateTime.Today.ToString("yyyyMMdd");
        }
    }
}
