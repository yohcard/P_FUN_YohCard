using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace P_Fun_YohCardis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            
            string filePath = @"C:\Users\ps16gnt\Documents\GitHub\P_FUN_YohCard\app\data\btc\bitcoin_2019-09-06_2024-09-04.csv";

            
            var lines = File.ReadAllLines(filePath).Skip(1);

            
            double[] dataX = new double[lines.Count()];
            double[] dataY = new double[lines.Count()];

            int index = 0;
            foreach (var line in lines)
            {
                var values = line.Split(',');
                dataX[index] = index + 1; 
                dataY[index] = double.Parse(values[2]); 
                index++;
            }

            // Plotting the data
            formsPlot1.Plot.Add.Scatter(dataX, dataY);
            formsPlot1.Refresh();
        }

        private void formsPlot1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
