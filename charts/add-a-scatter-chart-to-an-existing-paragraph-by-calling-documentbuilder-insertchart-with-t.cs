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

            // Add a paragraph that will later contain the scatter chart.
            builder.Writeln("Paragraph that will hold a scatter chart:");

            // Retrieve the paragraph we just added.
            var paragraph = doc.FirstSection.Body.Paragraphs[0];

            // Move the builder's cursor to the existing paragraph.
            builder.MoveTo(paragraph);

            // Insert a scatter chart into the paragraph using the InsertChart overload.
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 400, 300);
            Chart chart = chartShape.Chart;

            // Clear the default demo series.
            chart.Series.Clear();

            // Add a series with X and Y values.
            chart.Series.Add(
                "Sample Series",
                new double[] { 1.0, 2.0, 3.0, 4.0 },
                new double[] { 10.0, 20.0, 15.0, 30.0 });

            // Save the document to the working directory.
            doc.Save("ScatterChart.docx");
        }
    }
}
