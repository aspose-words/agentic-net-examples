using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartSeriesExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart with specific dimensions.
            Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = chartShape.Chart;

            // Clear any default series that were added automatically.
            chart.Series.Clear();

            // Add a labeled series in one step using the overload that accepts a name and values.
            chart.Series.Add("Quarterly Sales", new double[] { 15000, 20000, 25000, 30000 });

            // Save the document to a file.
            doc.Save("ChartSeriesLabeled.docx");
        }
    }
}
