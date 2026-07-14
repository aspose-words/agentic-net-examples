using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart related types

namespace ScatterChartExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an initial paragraph with some text.
            builder.Writeln("This paragraph will contain a scatter chart below.");

            // Move the builder to the first paragraph (index 0, node index 0).
            // This ensures the chart is inserted right after the paragraph.
            builder.MoveToParagraph(0, 0);

            // Insert a scatter chart using the overload that specifies chart type, width, and height.
            // Width and height are specified in points (1 point = 1/72 inch).
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 500, 300);
            Chart chart = chartShape.Chart;

            // Optional: clear the demo data that comes with the chart.
            chart.Series.Clear();

            // Add a series with X and Y values for the scatter plot.
            chart.Series.Add(
                "Sample Series",
                new double[] { 1.0, 2.5, 3.8, 5.0, 6.2 },
                new double[] { 2.0, 3.5, 1.8, 4.0, 5.5 });

            // Save the document to the working directory.
            doc.Save("ScatterChart.docx");
        }
    }
}
