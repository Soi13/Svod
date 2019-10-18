using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace Svod
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Excel.Application ObjExcel;
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = openFileDialog1.FileName;
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            /////Создание объекта БДДС подразделения
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["ПЛАН В"];
            /////////

            /////Создание объекта сводный БДДС
            Microsoft.Office.Interop.Excel.Application ObjExcel1 = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook1 = ObjExcel1.Workbooks.Open(openFileDialog2.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet1;
            ObjWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook1.Sheets["ПЛАН В"];
            /////////

            //////////////////Заполнение ПЛАН В/////////////////////////////
            int nn = 1;
            progressBar1.Maximum = 505;
            label5.Text = "Заполнение ПЛАН В";

            //Цикл по БДДС подразделения  
            for (int s = 4; s<=505; s++)
            {
                double stolbec_summa_bdds_podr = Convert.ToDouble(ObjWorkSheet.Cells[s, 5].value);

                //Определение цвета. Если не желтый, то пропускаем строку
                int color = int.Parse(ObjWorkSheet.Cells[s, 5].Interior.Color.ToString());
                if (color != 65535) { continue; }

                if (stolbec_summa_bdds_podr > 0)
                {
                    double stolbec_summa_svod = Convert.ToDouble(ObjWorkSheet1.Cells[s, 5].value);
                    double sum_common = stolbec_summa_svod + stolbec_summa_bdds_podr;
                    ObjWorkSheet1.Cells[s, 5] = sum_common;

                }
                
                for (int d = 15; d <= 45; d++)
                 {
                     //string date_bdds_podr = ObjWorkSheet.Cells[3, d].Text;
                     double summa_bdds_podr = Convert.ToDouble(ObjWorkSheet.Cells[s, d].value);

                     if (summa_bdds_podr>0)
                     {
                         double st = Convert.ToDouble(ObjWorkSheet1.Cells[s, d].value);
                         double sum = st + summa_bdds_podr;
                         ObjWorkSheet1.Cells[s, d] = sum;
                     }
         
                 }
                nn = nn + 1;
                progressBar1.Value = nn;
                label4.Text ="Обработано строк: " + Convert.ToString(nn);
            }
            ///////////////////////////////////////////////////////////////////


            
            
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["ПЛАН П"];
            /////////
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet3;
            ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook1.Sheets["ПЛАН П"];
            /////////

            //////////////////Заполнение ПЛАН П/////////////////////////////
            int nn1 = 1;
            progressBar1.Maximum = 505;
            label5.Text = "Заполнение ПЛАН П";

            //Цикл по БДДС подразделения  
            for (int s1 = 4; s1 <= 505; s1++)
            {
                if (s1 == 4 || s1 == 5 || s1 == 6 || s1 == 26 || s1 == 31 || s1 == 37 || s1 == 53 || s1 == 55 || s1 == 59 || s1 == 62 || s1 == 64 || s1 == 75 || s1 == 89 || s1 == 100 || s1 == 136 || s1 == 146 || s1 == 153 || s1 == 156 || s1 == 164 || s1 == 168 || s1 == 175 || s1 == 184 || s1 == 185 || s1 == 195 || s1 == 196 || s1 == 205 || s1 == 209 || s1 == 210 || s1 == 222) continue;
                
                double stolbec_summa_bdds_podr1 = Convert.ToDouble(ObjWorkSheet2.Cells[s1, 5].value);
                if (stolbec_summa_bdds_podr1 > 0)
                {
                    double stolbec_summa_svod1 = Convert.ToDouble(ObjWorkSheet3.Cells[s1, 5].value);
                    double sum_common1 = stolbec_summa_svod1 + stolbec_summa_bdds_podr1;
                    ObjWorkSheet3.Cells[s1, 5] = sum_common1;

                }
                         
                for (int d1 = 15; d1 <= 45; d1++)
                {
                    //string date_bdds_podr = ObjWorkSheet.Cells[3, d].Text;
                    double summa_bdds_podr1 = Convert.ToDouble(ObjWorkSheet2.Cells[s1, d1].value);

                    if (summa_bdds_podr1 > 0)
                    {
                        double st1 = Convert.ToDouble(ObjWorkSheet3.Cells[s1, d1].value);
                        double sum = st1 + summa_bdds_podr1;
                        ObjWorkSheet3.Cells[s1, d1] = sum;
                    }

                }
                nn1 = nn1 + 1;
                progressBar1.Value = nn1;
                label4.Text = "Обработано строк: " + Convert.ToString(nn1);
            }
            ///////////////////////////////////////////////////////////////////
            
            
            ObjWorkBook.Close();
            ObjExcel.Quit();
            ObjWorkBook = null;
            ObjWorkSheet = null;
            ObjExcel = null;

            ObjExcel1.Visible = true;
            
            GC.Collect(); 


        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {                
                    label3.Text = openFileDialog2.FileName;
            }
        }
    }
}
