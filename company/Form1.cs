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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cellhead a = new cellhead();
            a.Show();
        }
        homecompanyEntities1 db = new homecompanyEntities1();
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        public int changecheck = 0;
        private void Form1_Load(object sender, EventArgs e)
        {
            var dat = from n in db.TSell
                      orderby n.Date,n.OrderNo                     
                      select new
                      {
                          訂單代號 = n.OrderNo,
                          n.Po,
                          出貨日期 = n.Date,
                          狀況=n.Out,
                          數量 = n.Quantity,
                          規格 = n.Size,
                          商品代號 = n.ItemNo,
                          淨重 = n.NW,
                          總重 = n.GW,
                          港口 = n.Harber,
                          港口代號 = n.HarberCode,
                          成立日期 = n.CDate,
                      };
            int cxp = countPerPage * page;
            dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
           
            dataGridView1.Font = new Font("標楷體", 15);

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                DataGridViewColumn column = dataGridView1.Columns[i];
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }

            this.dataGridView1.SelectionMode =
            DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.MultiSelect = false;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;


            //dataGridView1.AutoSizeColumnsMode= DataGridViewAutoSizeColumnsMode.Fill;
            string[] y =
            {         "",      
                DateTime.Today.AddYears(1).Year.ToString(),
                DateTime.Today.Year.ToString(),              
                DateTime.Today.AddYears(-1).Year.ToString()
            };
            comboBox1.DataSource = y;
            comboBox1.SelectedItem = DateTime.Today.Year.ToString();

            comboBox2.SelectedItem = DateTime.Now.Month.ToString();

            comboBox3.SelectedItem = "成立日期";
            label3.Text = "第" + page.ToString() + "頁";

          
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //重新整理
            var dat = from n in db.TSell
                      orderby n.Date, n.OrderNo
                      select new
                      {
                          訂單代號 = n.OrderNo,
                          n.Po,
                          出貨日期 = n.Date,
                          狀況 = n.Out,
                          數量 = n.Quantity,
                          規格 = n.Size,
                          商品代號 = n.ItemNo,
                          淨重 = n.NW,
                          總重 = n.GW,
                          港口 = n.Harber,
                          港口代號 = n.HarberCode,
                          成立日期 = n.CDate,

                      };
             int cxp = countPerPage * page;
            dataGridView1.DataSource =  dat.Skip(cxp).Take(countPerPage).ToList();
            label3.Text = "第" + page.ToString() + "頁";
            page = 0;
            comboBox1.Text = " ";
            
            comboBox2.Text = " ";
            countPerPage = 10;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //年份搜尋
            var dat = from n in db.TSell
                      orderby n.Date, n.OrderNo
                      select new
                      {
                          訂單代號 = n.OrderNo,
                          n.Po,
                          出貨日期 = n.Date,
                          狀況 = n.Out,
                          數量 = n.Quantity,
                          規格 = n.Size,
                          商品代號 = n.ItemNo,
                          淨重 = n.NW,
                          總重 = n.GW,
                          港口 = n.Harber,
                          港口代號 = n.HarberCode,
                          成立日期 = n.CDate,
                      };
            int cxp = countPerPage * page;
            if(comboBox3.Text == "出貨日期")
            {                               
                dat = dat.Where(n => n.出貨日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {

                dat=dat.Where(n => n.成立日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            label3.Text = "第" + page.ToString() + "頁";
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //月份搜尋
            var dat = from n in db.TSell
                      orderby n.Date, n.OrderNo                     
                      select new
                      {
                          訂單代號 = n.OrderNo,
                          n.Po,
                          出貨日期 = n.Date,
                          狀況 = n.Out,
                          數量 = n.Quantity,
                          規格 = n.Size,
                          商品代號 = n.ItemNo,
                          淨重 = n.NW,
                          總重 = n.GW,
                          港口 = n.Harber,
                          港口代號 = n.HarberCode,
                          成立日期 = n.CDate,
                      };
            string tt = comboBox1.Text + comboBox2.Text;
            int cxp = countPerPage * page;
            if (comboBox3.Text == "訂單日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            label3.Text = "第" + page.ToString() + "頁";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            see s = new see();
            s.Show();
        }
        //datagridview點選顯示資料
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
           
            string value = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
           
            changecheck = 1;
            var n = (from m in db.TSell
                     where m.OrderNo == value
                     select m).FirstOrDefault();
            if (n != null)
            {
                comboBox5.Text = n.HarberCode;
                comboBox5.SelectedItem = n.HarberCode;
                textBox1.Text = n.ItemNo;
                textBox8.Text = n.Po;
                textBox7.Text = n.Date.ToString();
                textBox9.Text = n.CDate.ToString();
                textBox2.Text = n.Quantity.ToString();
                textBox3.Text = n.Size;
                textBox4.Text = n.NW;
                textBox5.Text = n.GW;
                textBox6.Text = n.Out;
                
                textBox11.Text = value;
                textBox10.Text = n.Al.ToString();
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        int page =0;
        int countPerPage = 10;

       

        private void button4_Click(object sender, EventArgs e)
        {          
            page += 1;
            var count = from n in db.TSell
                        group n by n.Id
                        into m select new
                        {
                            c = m.Count()
                        };

            int c = Convert.ToInt32(count.Count());
            int cxp = countPerPage * page;

            if (c- (countPerPage * page) < countPerPage)
            {
                countPerPage = c - (countPerPage * page);
                if (countPerPage < 0) countPerPage = 10;
                page--;
            };
            
            var dat = from n in db.TSell
                      orderby n.Date, n.OrderNo
                      select new
                      {
                          訂單代號 = n.OrderNo,
                          n.Po,
                          出貨日期 = n.Date,
                          狀況 = n.Out,
                          數量 = n.Quantity,
                          規格 = n.Size,
                          商品代號 = n.ItemNo,
                          淨重 = n.NW,
                          總重 = n.GW,
                          港口 = n.Harber,
                          港口代號 = n.HarberCode,
                          成立日期 = n.CDate,
                      };

            //抓讀取值
            string tt = comboBox1.Text + comboBox2.Text;
            if (comboBox3.Text == "出貨日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "訂單日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            countPerPage = 10;
            label3.Text = "第"+page.ToString()+"頁";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            page -= 1;
            if (page <= 0) page = 0;

            var dat = from n in db.TSell
                      orderby n.Date, n.OrderNo
                      select new
                      {
                          訂單代號 = n.OrderNo,
                          n.Po,
                          出貨日期 = n.Date,
                          狀況 = n.Out,
                          數量 = n.Quantity,
                          規格 = n.Size,
                          商品代號 = n.ItemNo,
                          淨重 = n.NW,
                          總重 = n.GW,
                          港口 = n.Harber,
                          港口代號 = n.HarberCode,
                          成立日期 = n.CDate,
                      };
            label3.Text = "第" + page.ToString() + "頁";

            //抓讀取值
            string tt = comboBox1.Text + comboBox2.Text;
            int cxp = countPerPage * page;
            if (comboBox3.Text == "出貨日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(comboBox1.Text));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
            if (comboBox3.Text == "訂單日期")
            {
                dat = dat.Where(n => n.出貨日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();

            }
            if (comboBox3.Text == "成立日期")
            {
                dat = dat.Where(n => n.成立日期.ToString().Contains(tt));
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            }
        }

        //修改
        private void button8_Click(object sender, EventArgs e)
        {
            if(changecheck == 1)
            {
                string search = textBox11.Text;
                var n1 = (from m in db.TSell
                         where m.OrderNo == search
                         select m).FirstOrDefault();
                n1.CDate = Convert.ToInt32(textBox9.Text);
                n1.Date = Convert.ToInt32(textBox7.Text);
                n1.Quantity = Convert.ToInt32(textBox2.Text);
                n1.ItemNo = textBox1.Text;
                n1.HarberCode = comboBox5.Text;
                n1.Size = textBox3.Text;
                n1.Po = textBox8.Text;
                n1.NW = textBox4.Text;
                n1.GW = textBox5.Text;
                n1.Out = textBox6.Text;
                n1.Al = Convert.ToInt32(textBox10.Text);
                db.SaveChanges();
                MessageBox.Show("修改成功");

                var dat = from n in db.TSell
                          orderby n.Date, n.OrderNo
                          select new
                          {
                              訂單代號 = n.OrderNo,
                              n.Po,
                              出貨日期 = n.Date,
                              狀況 = n.Out,
                              數量 = n.Quantity,
                              規格 = n.Size,
                              商品代號 = n.ItemNo,
                              淨重 = n.NW,
                              總重 = n.GW,
                              港口 = n.Harber,
                              港口代號 = n.HarberCode,
                              成立日期 = n.CDate,

                          };
                int cxp = countPerPage * page;
                dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
                label3.Text = "第" + page.ToString() + "頁";

                string Filestr = $@"C:\hw4\{textBox8.Text}";

                Excel.Application app = new Excel.Application();
                Excel.Workbook exwb = app.Workbooks.Add();

                Excel.Worksheet ws = new Excel.Worksheet();
                ws = exwb.Worksheets[1];
                ws.Name = "cell";

                if (comboBox5.Text == "66")
                {
                    app.Cells[3, 1] = "66";
                    app.Cells[5, 1] = "LAEM CHABANG";
                    app.Cells[8, 2] = textBox8.Text;
                    app.Cells[19, 1] = "66";
                    app.Cells[21, 1] = "LAEM CHABANG";
                    app.Cells[24, 2] = textBox8.Text;

                }
                else if (comboBox5.Text == "84")
                {
                    app.Cells[3, 1] = "84";
                    app.Cells[5, 1] = "HO CHI MINH";
                    app.Cells[8, 2] = textBox8.Text;
                    app.Cells[19, 1] = "84";
                    app.Cells[21, 1] = "HO CHI MINH";
                    app.Cells[24, 2] = textBox8.Text;

                }
                else if (comboBox5.Text == "SMART TECH")
                {
                    app.Cells[3, 1] = "SMART TECH";
                    app.Cells[5, 1] = "BAVET";
                    app.Cells[8, 2] = textBox8.Text;
                    app.Cells[19, 1] = "SMART TECH";
                    app.Cells[21, 1] = "BAVET";
                    app.Cells[24, 2] = textBox8.Text;
                }


                //右上
                app.Cells[4, 1] = "--------";

                app.Cells[6, 1] = "MADE IN TAIWAN";
                app.Cells[7, 1] = "R.O.C.";
                app.Cells[8, 1] = @"PO:";
                app.Cells[9, 1] = @"C/No.:";
                app.Cells[11, 5] = textBox11.Text;
                app.Cells[11, 1] = "A";

                app.Cells[4, 5] = "ARTCLE:STEEL PIN";
                app.Cells[5, 5] = "ITEM NO. ";
                app.Cells[5, 6] = textBox1.Text;
                app.Cells[6, 5] = "SIZE ";
                app.Cells[6, 6] = textBox2.Text;
                app.Cells[7, 5] = "QNTY ";
                app.Cells[7, 6] = textBox3.Text ;
                app.Cells[7, 7] = "PC";
                app.Cells[8, 5] = "N.W";
                app.Cells[8, 6] = textBox4.Text;
                app.Cells[8, 7] =   "KG";
                app.Cells[9, 5] = "G.W";
                app.Cells[9, 6] = textBox5.Text ;
                app.Cells[9, 7] =  " KG";
                app.Cells[3, 5] = textBox7.Text;  //日期
               
                
                //右下

                app.Cells[20, 1] = "--------";
                app.Cells[22, 1] = "MADE IN TAIWAN";
                app.Cells[23, 1] = "R.O.C.";
                app.Cells[24, 1] = @"PO:";
                app.Cells[25, 1] = @"C/No.:";
                app.Cells[27, 5] = textBox11.Text;
                app.Cells[27, 1] = "B";

                app.Cells[20, 5] = "ARTCLE:STEEL PIN";
                app.Cells[21, 5] = "ITEM NO. ";
                app.Cells[21, 6] = textBox1.Text;
                app.Cells[22, 5] = "SIZE ";
                app.Cells[22, 6] = textBox2.Text;
                app.Cells[23, 5] = "QNTY ";
                app.Cells[23, 6] = textBox3.Text ;
                app.Cells[23, 7] =  "PC";
                app.Cells[24, 5] = "N.W";
                app.Cells[24, 6] = textBox4.Text;
                app.Cells[24, 7] = "KG";
                app.Cells[25, 5] = "G.W";
                app.Cells[25, 6] = textBox5.Text;
                app.Cells[25, 7] =  "KG";
                app.Cells[19, 5] = textBox7.Text;
                
                


                exwb.SaveAs(Filestr);

             


                ws = null;
                exwb.Close();
                exwb = null;
                app.Quit();
                app = null;

            }
            else
            {
                MessageBox.Show("請先選擇資料");
            }
        }

        //列印
        private void button6_Click(object sender, EventArgs e)
        {
            if(changecheck == 1) 
            {
                Excel.Application app = new Excel.Application();
                string Filestr = $@"C:\hw4\{textBox8.Text}";
                app.Visible = true;
                app.Workbooks.Open(Filestr);
            }
            else
            {
                MessageBox.Show("請先選擇資料");
            }
           
        }
        //刪除
        private void button7_Click(object sender, EventArgs e)
        {
            if(changecheck == 1)
            {
                string search = textBox11.Text;
            var n = (from m in db.TSell
                     where m.OrderNo == search
                     select m).FirstOrDefault();
            db.TSell.Remove(n);
            db.SaveChanges();
            MessageBox.Show("刪除成功");
                textBox9.Text="";
                textBox7.Text = "";
                textBox2.Text = "";
                textBox1.Text = "";
                comboBox5.Text = "";
                textBox3.Text = "";
                textBox8.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox10.Text = "";
            }
            else
            {
                MessageBox.Show("請先選擇資料");
            }

            var dat = from n in db.TSell
                      orderby n.Date, n.OrderNo
                      select new
                      {
                          訂單代號 = n.OrderNo,
                          n.Po,
                          出貨日期 = n.Date,
                          狀況 = n.Out,
                          數量 = n.Quantity,
                          規格 = n.Size,
                          商品代號 = n.ItemNo,
                          淨重 = n.NW,
                          總重 = n.GW,
                          港口 = n.Harber,
                          港口代號 = n.HarberCode,
                          成立日期 = n.CDate,

                      };
            int cxp = countPerPage * page;
            dataGridView1.DataSource = dat.Skip(cxp).Take(countPerPage).ToList();
            label3.Text = "第" + page.ToString() + "頁";
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }
        //搜尋
        private void searchB_Click(object sender, EventArgs e)
        {
            if (textBox11.Text == "")
                MessageBox.Show("請輸入訂單代號");
            
            var n = (from m in db.TSell
                    where m.OrderNo.Contains(textBox11.Text)
                    select m).FirstOrDefault();
            if (n != null)
            {
                comboBox5.Text = n.HarberCode;
                comboBox5.SelectedItem = n.HarberCode;
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
            else
            {
                MessageBox.Show("查無此項");
            }



        }
    }
}
