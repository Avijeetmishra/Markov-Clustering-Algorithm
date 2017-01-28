using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using Microsoft.Office.Interop.Excel;
using System.Windows.Media;

namespace Markov_Cluster_Algorithm
{
    public partial class Form1 : Form
    {
        string[,] data = new string[400, 400];
        int count = 1;
        int power = 0;
        int inflate = 0;
        Dictionary<string, int> nodes = new Dictionary<string, int>();
        int itr = 0;
        public Form1()
        {
            InitializeComponent();
        }
        private double[,] expand_inflate(double[,] final, double[,] test)
        {
            double[,] final1 = new double[count, count];
            //expand
            for (int i = 1; i < count; i++)
            {
                for (int j = 1; j < count; j++)
                {
                    final1[i, j] = 0;
                    for (int k = 1; k < count; k++)
                    {
                        final1[i, j] += final[i, k] * final[k, j];

                    }
                }
            }
            for (int i = 1; i < count; i++)
            {
                for (int j = 1; j < count; j++)
                {
                    final[i, j] = final1[i, j];
                }
            }

            int p = power;
            while (p > 2)
            {
                double[,] temp = new double[count, count];
                for (int i = 1; i < count; i++)
                {
                    for (int j = 1; j < count; j++)
                    {
                        temp[i, j] = 0;
                        for (int k = 1; k < count; k++)
                        {
                            temp[i, j] += final1[i, k] * final1[k, j];
                        }
                    }
                }
                for (int i = 1; i < count; i++)
                {
                    for (int j = 1; j < count; j++)
                    {
                        final1[i, j] = temp[i, j];
                        final[i, j] = temp[i, j];
                    }
                }
                p--;
            }

            // inflate
            int f = inflate;
            while (f != 1)
            {
                for (int i = 1; i < count; i++)
                {
                    for (int j = 1; j < count; j++)
                    {
                        final[i, j] = final1[i, j] * final[i, j];
                    }
                }
                f--;
            }

            //normalize
            for (int i = 1; i < count; i++)
            {
                double sum = 0;
                for (int j = 1; j < count; j++)
                {
                    sum = sum + final[j, i];
                }
                for (int j = 1; j < count; j++)
                {
                    final[j, i] = final[j, i] / sum;
                  
                }
            }
            //Check for Convergence
            int t = 0;
            for (int i = 1; i < count; i++)
            {

                for (int j = 1; j < count; j++)
                {
                    if (final[i, j] == test[i, j])
                    {
                        t++;
                    }
                    else
                    {
                        test[i, j] = final[i, j];
                    }
                }
            }
            itr++;
            if (t != ((count - 1) * (count - 1)))
            {
                final = expand_inflate(final, test);
            }
            return final;  

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filepath = null;
            
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                filepath = openFileDialog1.FileName;
            }

            inflate = Convert.ToInt32(textBox1.Text);
            power = Convert.ToInt32(textBox2.Text);
            Microsoft.Office.Interop.Excel.Application IExcel = new Microsoft.Office.Interop.Excel.Application();
            string fileName = filepath;
            //open the workbook
            Workbook workbook = IExcel.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            //select the first sheet        
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            //find the used range in worksheet
            Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

            //get an object array of all of the cells in the worksheet (their values)
            object[,] valueArray = (object[,])excelRange.get_Value(
                        XlRangeValueDataType.xlRangeValueDefault);
            string name = workbook.Name;
            int[] val=new int[2];
            string[,] inp = new string[worksheet.UsedRange.Rows.Count, worksheet.UsedRange.Columns.Count];
            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
            {
                for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                {
                    if (!nodes.ContainsKey(Convert.ToString(valueArray[row, col])))
                    {
                        nodes.Add(Convert.ToString(valueArray[row, col]), count);
                        val[col - 1] = count;
                        count++;
                        data[val[col - 1], 0] = Convert.ToString(valueArray[row, col]);
                        data[0, val[col - 1]] = Convert.ToString(valueArray[row, col]);
                        //inp[row - 1, col - 1] = Convert.ToString(valueArray[row, col]);
                    }
                    else
                    {
                        nodes.TryGetValue(Convert.ToString(valueArray[row, col]), out val[col-1]);
                       
                    }
                    //access each cell
                    
                    
                }

                data[val[0], val[1]] = Convert.ToString(1);
                data[val[1], val[0]] = Convert.ToString(1);
            }
            var watch = System.Diagnostics.Stopwatch.StartNew();
            double[,] final = new double[count, count];
            double[,] test = new double[count, count];
            double[,] final1 = new double[count, count];
            //self loop
            for (int i = 1; i < count; i++)
            {
                for (int j = 1; j < count; j++)
                {
                    if (i == j)
                    {
                        final[i, j] = 1;
                    }
                    else if (data[i, j] == null)
                    {
                        final[i, j] = 0;
                    }
                    else
                    {
                        final[i, j] = Convert.ToDouble(data[i, j]);
                    } 
                }
            }
            //normalize the matrix
            for (int i = 1; i < count; i++)
            {
                double sum = 0;
                for (int j = 1; j < count; j++)
                {
                    sum = sum + final[j, i];
                }
                for (int j = 1; j < count; j++)
                {
                    final[j, i] = final[j, i]/sum;
                }
            }
            for (int i = 1; i < count; i++)
            {
                for (int j = 1; j < count; j++)
                {
                    test[i, j] = final[i, j];
                }
            }

            final =expand_inflate(final, test);
            //Cluster Discovery
            string[] clu = new string[count];
            int cluster = 1;
            for (int i = 1; i < count; i++)
            {
                int flag = 0;
                double temp = 999;
                int temp_flag = 0;
                for (int j = 1; j < count; j++)
                {
                    if (final[i, j] != 0 && temp_flag == 0)
                    {
                        temp = final[i, j];
                        temp_flag = 1;
                    }
                    if (final[i, j] == temp)
                    {
                        if (clu[j] == null)
                        {
                            clu[j] = cluster.ToString();
                            flag = 1;
                        }
                        
                    }
                }
                if (flag == 1)
                {
                    cluster++;
                }
            }
            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            label3.Text = label3.Text + (cluster-1).ToString();
            label4.Text = label4.Text + (itr).ToString();
            label5.Text = label5.Text + (elapsedMs).ToString() + " Milliseconds";

            clu[0] = "*Vertices " + (count - 1).ToString();
            System.IO.File.WriteAllLines(@filepath + "power_" + power + "inflate_" + inflate + ".clu", clu);
            //r++;
           
           /* for (int i = 0; i < count; i++)
            {
                for (int j = 0; j <count; j++)
                {
                    Console.Write(final[i, j] + "\t");
                }
                Console.WriteLine();
            }*/
       
    }
    }
}
