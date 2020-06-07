using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelDataReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataSet sonuc;
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Dosyası |*.xls;*.xlsx;*.xlsm;*.csv"; ;
            var sonuc = dialog.ShowDialog();

            if (sonuc == DialogResult.OK)
            {
                textBox1.Text = dialog.FileName;
            }
            dosyaoku(0);
           
        }
        private void dosyaoku(int sayfa)
        {
            if (textBox1.Text != string.Empty)
            {
                string yol = textBox1.Text;
                //UZANTIYI ALIR VE UZANTISINI KÜÇÜK HAFRE ÇEVİRİR.
                var uzanti = Path.GetExtension(textBox1.Text).ToLower();
                FileStream stream = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                IExcelDataReader excelDataReader = null;
                if (uzanti == ".xlsx")
                {
                    excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                else if (uzanti == ".xls") //.xls uzantısı için
                {
                    excelDataReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (uzanti == ".csv")
                {
                    excelDataReader = ExcelReaderFactory.CreateCsvReader(stream);
                }
                if (excelDataReader == null)
                    return;

                sonuc = excelDataReader.AsDataSet();                        
                dataGridView1.DataSource = sonuc.Tables[sayfa];
                var deger = sonuc.Tables[sayfa].Rows[0];
                excelDataReader.Close();

                for (int i = 0; i < deger.ItemArray.Length; i++)
                {
                    dataGridView1.Columns[i].HeaderText = deger[i].ToString();
                }
                dataGridView1.Rows.RemoveAt(0);

            }
            else
            {
                MessageBox.Show("Bağlantı yolunu seçiniz", "Bağlantı Yolu Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //datagriddeki verilere erişmek istenilirse 
            for(int i=0; i<dataGridView1.Rows.Count-1; i++)
            {
               double x=(double)dataGridView1.Rows[i].Cells[2].Value; //x kordinatını 
               double y=(double)dataGridView1.Rows[i].Cells[3].Value; //y kordinatını 
                
            }
        }
    }
}
