using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;       // Chart related classes

namespace AsposeWordsChartsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a paragraph that will precede the chart.
            builder.Writeln("Scatter chart inserted below this paragraph:");

            // Insert a scatter chart at the current builder position using the overload that specifies type, width and height.
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 500, 300);
            Chart chart = chartShape.Chart;

            // Remove the default demo data and add a custom series.
            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new[] { 1.0, 2.5, 4.0, 5.5 },   // X‑values
                new[] { 3.0, 4.5, 2.0, 6.0 }    // Y‑values
            );

            // Save the document to the local file system.
            doc.Save("ScatterChartExample.docx");
        }
    }
}
