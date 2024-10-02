﻿using ScottPlot;
using ScottPlot.WinForms;

namespace TES_FUN2.Tests
{
    [TestClass()]
    public class Form1Tests
    {
        [TestMethod()]

        // Vérifie que le plot n'est pas null lorsque l'on affiche les datas avec la methode plotdata
        // check if the plot is nul when we display data with the function plot data
        public void PlotDataTest()
        {   


            //arrange
            var plot = new Form1();
            plot.FormsPlot = new FormsPlot();

            var data = new Dictionary<string, Dictionary<DateTime, double>>
            {
                { "btc", new Dictionary<DateTime, double> { { DateTime.Now, 45000 } } },
                { "sol", new Dictionary<DateTime, double> { { DateTime.Now, 8900 } } }
            };

            // Act
            plot.PlotData(data);

            // Assert
            Assert.IsNotNull(plot.FormsPlot.Plot); 
        }
        [TestMethod()]

        //verifie que le nombre de 
        public void PlotDataTest2()
        {
            //arrange
            var plot = new Form1();
            plot.FormsPlot = new FormsPlot();

            var data = new Dictionary<string, Dictionary<DateTime, double>>
            {
                { "btc", new Dictionary<DateTime, double> { { DateTime.Now, 45000 } } },
                { "sol", new Dictionary<DateTime, double> { { DateTime.Now, 8900 } } }
            };

            // Act
            plot.PlotData(data);

            // Assert
            Assert.AreEqual(2, plot.FormsPlot.Plot.GetPlottables().Count(), "Le nombre de tracés ajoutés est incorrect.");
        }
    }
}