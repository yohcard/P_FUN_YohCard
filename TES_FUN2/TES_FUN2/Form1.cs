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
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"bitcoin_2019-09-06_2024-09-04(2).xlsx");

            // Dictionnaire pour stocker les données (DateTime pour les dates, double pour les prix)
            Dictionary<DateTime, double> data = new Dictionary<DateTime, double>();

            // Lecture du fichier Excel
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new XSSFWorkbook(stream);
                var sheet = workbook.GetSheetAt(0); // Obtenir la première feuille

                // Lire les lignes (en ignorant la première qui est l'en-tête)
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row != null)
                    {
                        var dateCell = row.GetCell(1);
                        var priceCell = row.GetCell(5);

                        // Vérifier si les cellules sont non nulles
                        if (dateCell != null && priceCell != null)
                        {


                               
                                DateTime closeDate = Convert.ToDateTime(dateCell.DateCellValue); // Récupérer la 2e colonne (date de fermeture)
                            double closePrice = priceCell.CellType == NPOI.SS.UserModel.CellType.Numeric ? priceCell.NumericCellValue : 0; // Récupérer la 6e colonne (prix de clôture)
                                data.Add(closeDate, closePrice);
                            
                        }
                    }
                }
            }

            // Préparation des données pour ScottPlot
            DateTime[] dataX = data.Keys.ToArray();  // Utiliser DateTime pour les X (dates)
            double[] dataY = data.Values.ToArray();  // Utiliser les prix pour les Y

            // Tracé du graphique
            formsPlot1.Plot.XLabel("Date");
            formsPlot1.Plot.YLabel("Prix");
            formsPlot1.Plot.Add.Scatter(dataX, dataY);

            // Affichage des ticks pour les dates
            formsPlot1.Plot.Axes.DateTimeTicksBottom();

            LegendItem item1 = new()
            {
                LineColor = ScottPlot.Color.FromHex("C43E1C"),
                MarkerColor = ScottPlot.Color.FromHex("C43E1C"),
                Label = "btc"
            };

            LegendItem[] items = { item1 };
            formsPlot1.Plot.ShowLegend(items);
            formsPlot1.Refresh();
        }
    }
}
