using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart with the ShadedPlot style to ensure consistent branding.
            Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300, ChartStyle.ShadedPlot);
            Chart chart = chartShape.Chart;

            // Replace the demo data with custom series.
            chart.Series.Clear();
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            chart.Series.Add("Sales", categories, new double[] { 15000, 21000, 18000, 24000 });

            // Save the document.
            doc.Save("ChartWithStyle.docx");
        }
    }
}
