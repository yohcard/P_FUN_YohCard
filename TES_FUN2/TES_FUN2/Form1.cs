using ScottPlot;
using NPOI.XSSF.UserModel; // Pour lire le fichier Excel
using System.IO;

namespace TES_FUN2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CreateChart();
        }

        private void CreateChart()
        {
            // Chemins des fichiers
            string btcFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"bitcoin_2019-09-13_2024-09-11.xlsx");
            string solFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"solana_2019-09-13_2024-09-11.xlsx");
            string ethFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"ethereum_2019-09-13_2024-09-11.xlsx");

            // Dictionnaires pour stocker les donn�es
            Dictionary<DateTime, double> btcData = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> solData = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> ethData = new Dictionary<DateTime, double>();


            // Fonction pour lire les donn�es � partir d'un fichier Excel
            void ReadExcelData(string filePath, Dictionary<DateTime, double> data)
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new XSSFWorkbook(stream);
                    var sheet = workbook.GetSheetAt(0);

                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        var row = sheet.GetRow(i);
                        if (row != null)
                        {
                            var dateCell = row.GetCell(1);
                            var priceCell = row.GetCell(5);

                            // V�rification des cellules non nulles
                            if (dateCell != null && priceCell != null)
                            {
                                DateTime closeDate = Convert.ToDateTime(dateCell.DateCellValue);

                                // V�rification que la date est valide et non d�j� pr�sente
                                if (closeDate != DateTime.MinValue && !data.ContainsKey(closeDate))
                                {
                                    double closePrice = priceCell.CellType == NPOI.SS.UserModel.CellType.Numeric ? priceCell.NumericCellValue : 0;
                                    data.Add(closeDate, closePrice);
                                }
                            }
                        }
                    }
                }
            }
            

            // Lecture des donn�es Bitcoin
            ReadExcelData(btcFilePath, btcData);

            // Lecture des donn�es Solana
            ReadExcelData(solFilePath, solData);

            // Lecture des donn�es etherum
            ReadExcelData(ethFilePath, ethData);
            

            // Pr�paration des donn�es pour ScottPlot (BTC)
            DateTime[] btcX = btcData.Keys.ToArray();
            double[] btcY = btcData.Values.ToArray();

            // Pr�paration des donn�es pour ScottPlot (Solana)
            DateTime[] solX = solData.Keys.ToArray();
            double[] solY = solData.Values.ToArray();

            // Pr�paration des donn�es pour ScottPlot (Solana)
            DateTime[] etlX = ethData.Keys.ToArray();
            double[] ethY = ethData.Values.ToArray();

            //gestion des labels du graphique
            formsPlot1.Plot.XLabel("Date");
            formsPlot1.Plot.YLabel("Prix");


            // Trac� du graphique Bitcoin
            formsPlot1.Plot.Add.Scatter(btcX, btcY, color: ScottPlot.Color.FromHex("C43E1C"));

            // Trac� du graphique Solana
            formsPlot1.Plot.Add.Scatter(solX, solY, color: ScottPlot.Color.FromHex("1C72C4"));

            // Trac� du graphique Solana
            formsPlot1.Plot.Add.Scatter(etlX, ethY, color: ScottPlot.Color.FromHex("000000"));

            // Affichage des ticks pour les dates
            formsPlot1.Plot.Axes.DateTimeTicksBottom();

            // Affichage de la l�gende
            formsPlot1.Plot.ShowLegend();

            // Rafra�chissement du graphique
            formsPlot1.Refresh();
        }

    }
}
