using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;        // Chart types and styles

namespace ChartStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for inserting content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart with the predefined ShadedPlot style.
            // Width and height are specified in points.
            Shape chartShape = builder.InsertChart(ChartType.Column, 400, 250, ChartStyle.ShadedPlot);
            Chart chart = chartShape.Chart;

            // Optional: set a title for the chart.
            chart.Title.Text = "Sales Overview";
            chart.Title.Show = true;

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document containing the styled chart.
            string outputPath = Path.Combine(outputDir, "ChartWithStyle.docx");
            doc.Save(outputPath);
        }
    }
}
