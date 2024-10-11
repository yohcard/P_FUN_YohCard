using ScottPlot;
using NPOI.XSSF.UserModel; 
using System.IO;
using NPOI.SS.Formula.Functions;
using ScottPlot.WinForms;

namespace TES_FUN2
{
    public partial class Form1 : Form
    {
        public FormsPlot FormsPlot
        {
            get { return formsPlot1; }
            set { formsPlot1 = value; }
        }

        // Liste des devises
        // Currencies list
        string[] currencies;
        bool yoyo;


        // Dictionnaires pour stocker les données et les chemins de fichiers
        // Dictionnary wich stock curenncies data and files path
        Dictionary<string, Dictionary<DateTime, double>> currencyData;
        Dictionary<string, string> filePaths;

        public Form1()
        {
            InitializeComponent();

            // Initialisation des monnaies
            // Initialize currencies
            currencies = new[] { "btc", "sol", "eth" };



            // Initialisation des données de devises et des chemins de fichiers
            // Initialize currencies data and files path
            currencyData = currencies.ToDictionary(c => c, c => new Dictionary<DateTime, double>());
            filePaths = currencies.ToDictionary(c => c, c => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{c}_2019-09-13_2024-09-11.xlsx"));
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            CreateChart();
        }

        private void CreateChart()
        {
            // Lecture des données
            // Reading data
            foreach (var currency in currencyData.Keys)
            {

                ReadExcelData(filePaths[currency], currencyData[currency]);

            }

            // Afficher le graphique complet au chargement
            // Show full graph on load
            PlotData(currencyData);
        }

        public void PlotData(Dictionary<string, Dictionary<DateTime, double>> data)
        {
            formsPlot1.Reset();

            // Définir les couleurs pour chaque monnaie
            // Set colors for each currency
            var colors = new Dictionary<string, string>
            {
                { currencies[0], "C43E1C" },
                { currencies[1], "1C72C4" },
                { currencies[2], "000000" }
            };

            foreach (var currency in data.Keys)
            {
                if (data[currency].Count > 2)
                {
                    var result = AvgMinMaxValue(data[currency]);
                    DateTime[] xValues = data[currency].Keys.ToArray();
                    double[] yValues = data[currency].Values.ToArray();
                    var scatter = formsPlot1.Plot.Add.Scatter(xValues, yValues, color: ScottPlot.Color.FromHex(colors[currency]));
                    scatter.LegendText = currency + ", avg : " + result.Item1 + ",  min : " + result.Item2 + ",  max : " + result.Item3;
                }

            }

            // Affichage des ticks pour les dates
            // Ticks display for dates
            formsPlot1.Plot.Axes.DateTimeTicksBottom();
            formsPlot1.Plot.ShowLegend(ScottPlot.Edge.Bottom);


            // Gestion des labels du graphique
            // Management of chart labels
            formsPlot1.Plot.YLabel("Prix de fermeture en USD");
            formsPlot1.Plot.Title("Plot that line !");
            formsPlot1.Plot.XLabel("Date");

            // Rafraîchissement du graphique
            // Refresh the chart
            formsPlot1.Refresh();
        }

        // Récupère le prix de fermeture et la date de chaque ligne pour chaque fichier
        // Retrieves the closing price and date of each line for each file
        public void ReadExcelData(string filePath, Dictionary<DateTime, double> data)
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


                //à refactoriser
                foreach (var row in rows)
                {
                    data.Add(row.Date, row.Price);
                }
            }
        }



        public (double, double, double) AvgMinMaxValue(Dictionary<DateTime, double> data)
        {

            var minValue = Math.Round(data.Values.Min(), 2);
            var maxValue = Math.Round(data.Values.Max(), 2);
            var avgValue = Math.Round(data.Values.Average(), 2);

            return (minValue, maxValue, avgValue);
        }

        private void DateFilterBtn_Click(object sender, EventArgs e)
        {
            // Récupérer les dates choisies dans les DateTimePickers
            // Retrieve dates selected in DateTimePickers
            DateTime startDate = dateTimePicker1.Value;
            DateTime endDate = dateTimePicker2.Value;

            // Pour chaque monnaie, création d'un dictionnaire contenant uniquement les données choisies
            // For each currency, creation of a dictionary containing only the selected data
            var filteredData = currencyData.ToDictionary(
                currency => currency.Key,
                currency => currency.Value.Where(d => d.Key >= startDate && d.Key <= endDate)
                .ToDictionary(d => d.Key, d => d.Value)
            );

            foreach (var row in filteredData.Values)
            {

                yoyo = row.Count == 0;
            }
            
            
            if (yoyo == true)
            {
                {
                    string NoDataforThisPeriod = "il n'y pas de données pour la periode choisie.";
                    MessageBox.Show(NoDataforThisPeriod);
                }

                // Utilisation de PlotData pour exposer les données filtrées
                // Using PlotData to expose filtered data

            }
            else
            {
                PlotData(filteredData);
            }
            }

            private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
            {

            }

            private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
            {

            }

            private void formsPlot1_Load(object sender, EventArgs e)
            {

            }

            private void resetBtn_Click(object sender, EventArgs e)
            {
                formsPlot1.Plot.Clear();
                CreateChart();
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