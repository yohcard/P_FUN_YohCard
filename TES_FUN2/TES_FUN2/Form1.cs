using ScottPlot;
using NPOI.XSSF.UserModel; 
using System.IO;
using NPOI.SS.Formula.Functions;

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

            // Dictionnaires pour stocker les données
            Dictionary<DateTime, double> btcData = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> solData = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> ethData = new Dictionary<DateTime, double>();


            // Fonction pour lire les données à partir d'un fichier Excel
            /*
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

                            // Vérification des cellules non nulles
                            if (dateCell != null && priceCell != null)
                            {
                                DateTime closeDate = Convert.ToDateTime(dateCell.DateCellValue);

                                // Vérification que la date est valide et non déjà présente
                                if (closeDate != DateTime.MinValue && !data.ContainsKey(closeDate))
                                {
                                    double closePrice = priceCell.CellType == NPOI.SS.UserModel.CellType.Numeric ? priceCell.NumericCellValue : 0;
                                    data.Add(closeDate, closePrice);
                                }
                            }
                        }
                    }
                }
            }*/

            //version linq
            void ReadExcelData(string filePath, Dictionary<DateTime, double> data)
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new XSSFWorkbook(stream);
                    var sheet = workbook.GetSheetAt(0);

                    var rows = Enumerable.Range(1, sheet.LastRowNum)
                                         .Select(i => sheet.GetRow(i))
                                         .Where(row => row != null)
                                         .Select(row => new
                                         {
                                             DateCell = row.GetCell(1),
                                             PriceCell = row.GetCell(5)
                                         })
                                         .Where(cells => cells.DateCell != null && cells.PriceCell != null)
                                         .Select(cells => new
                                         {
                                             Date = Convert.ToDateTime(cells.DateCell.DateCellValue),
                                             Price = cells.PriceCell.CellType == NPOI.SS.UserModel.CellType.Numeric ? cells.PriceCell.NumericCellValue : 0
                                         })
                                         .Where(item => item.Date != DateTime.MinValue && !data.ContainsKey(item.Date));

                    foreach (var row in rows)
                    {
                        data.Add(row.Date, row.Price);
                    }
                }
            }

            // Lecture des données Bitcoin
            ReadExcelData(btcFilePath, btcData);

            // Lecture des données Solana
            ReadExcelData(solFilePath, solData);

            // Lecture des données etherum
            ReadExcelData(ethFilePath, ethData);
           
            // Préparation des données pour ScottPlot (BTC)
            DateTime[] btcX = btcData.Keys.ToArray();
            double[] btcY = btcData.Values.ToArray();

            // Préparation des données pour ScottPlot (Solana)
            DateTime[] solX = solData.Keys.ToArray();
            double[] solY = solData.Values.ToArray();

            // Préparation des données pour ScottPlot (etherum)
            DateTime[] etlX = ethData.Keys.ToArray();
            double[] ethY = ethData.Values.ToArray();

            // Tracé du graphique Bitcoin
            var btc = formsPlot1.Plot.Add.Scatter(btcX, btcY, color: ScottPlot.Color.FromHex("C43E1C"));
            btc.LegendText = "btc";

            // Tracé du graphique Solana
            var sol = formsPlot1.Plot.Add.Scatter(solX, solY, color: ScottPlot.Color.FromHex("1C72C4"));
            sol.LegendText = "sol";

            // Tracé du graphique Solana
            var eth = formsPlot1.Plot.Add.Scatter(etlX, ethY, color: ScottPlot.Color.FromHex("000000"));
            eth.LegendText = "eth";

            // Affichage des ticks pour les dates
            formsPlot1.Plot.Axes.DateTimeTicksBottom();

            //gestion des labels du graphique
            formsPlot1.Plot.YLabel("Prix");
            formsPlot1.Plot.Title("Plot that line !");
            formsPlot1.Plot.XLabel("Date");

            
            // Rafraîchissement du graphique
            formsPlot1.Refresh();
        }
    }
}
