using ScottPlot;
using NPOI.XSSF.UserModel; 
using System.IO;
using NPOI.SS.Formula.Functions;

namespace TES_FUN2
{
    public partial class Form1 : Form
    {
        // Dictionnaires pour stocker les données
        Dictionary<DateTime, double> btcData = new Dictionary<DateTime, double>();
        Dictionary<DateTime, double> solData = new Dictionary<DateTime, double>();
        Dictionary<DateTime, double> ethData = new Dictionary<DateTime, double>();

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

            // Lecture des données
            ReadExcelData(btcFilePath, btcData);
            ReadExcelData(solFilePath, solData);
            ReadExcelData(ethFilePath, ethData);

            // Afficher le graphique complet au chargement
            PlotData(btcData, solData, ethData);
        }


        private void PlotData(Dictionary<DateTime, double> btcData, Dictionary<DateTime, double> solData, Dictionary<DateTime, double> ethData)
        {
            formsPlot1.Reset();
           
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

        //récupère le prix de fermeture et la date de chaque ligne pour chaque ligne du fichier
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void DateFilterBtn_Click(object sender, EventArgs e)
        {
            //récuperère les dates choisi dans les dateTimePickers
            DateTime startDate = dateTimePicker1.Value;
            DateTime endDate = dateTimePicker2.Value;

            //Pour chaque monnaie, création d'un dictionnaire contenant uniquement le données choisie
            var filteredBtcData = btcData.Where(d => d.Key >= startDate && d.Key <= endDate).ToDictionary(d => d.Key, d => d.Value);
            var filteredsolData = solData.Where(d => d.Key >= startDate && d.Key <= endDate).ToDictionary(d => d.Key, d => d.Value);
            var filteredethData = ethData.Where(d => d.Key >= startDate && d.Key <= endDate).ToDictionary(d => d.Key, d => d.Value);

            //utilisation de PlotData pour exposer les données filtré
            PlotData(filteredBtcData, filteredsolData, filteredethData);

        }
    }
}








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