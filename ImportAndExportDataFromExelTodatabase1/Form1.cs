using ExcelDataReader;
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

namespace ImportAndExportDataFromExelTodatabase1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] table = comboBox1.SelectedItem.ToString().Split(' ');
            DataTable dt = tableCollection[table[0]];
            dataGridView1.DataSource = dt;
        }

        DataTableCollection tableCollection;

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog=new OpenFileDialog() { Filter= "All files (*.*)|*.*" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = openFileDialog.FileName;
                    using(var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using(IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable=(_)=>new ExcelDataTableConfiguration() { UseHeaderRow=true}
                            });
                            tableCollection = result.Tables;
                            comboBox1.Items.Clear();
                            foreach(DataTable table in tableCollection)
                            {
                                comboBox1.Items.Add(table.TableName);
                            }
                        }
                    }
                }
            }
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            dataGridView1.Height = this.Height * 80 / 100;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string mahsulotturi = comboBox1.SelectedItem.ToString();
        }
    }
}
