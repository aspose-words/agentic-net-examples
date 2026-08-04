using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace ChartInTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a fixed column width so the cell size is deterministic.
            const double columnWidth = 300.0; // points
            builder.CellFormat.Width = columnWidth;

            // Start a table with two columns.
            builder.StartTable();

            // First cell – just some placeholder text.
            builder.InsertCell();
            builder.Write("Sample data");

            // Second cell – insert the chart.
            builder.InsertCell();

            // Insert a column chart. Width and height are set to match the cell width.
            Shape chartShape = builder.InsertChart(ChartType.Column, columnWidth, columnWidth * 0.6);
            if (!chartShape.HasChart)
                throw new InvalidOperationException("The inserted shape does not contain a chart.");

            // Ensure the chart scales proportionally with the cell dimensions.
            chartShape.Width = columnWidth;
            chartShape.Height = columnWidth * 0.6; // maintain the same aspect ratio used on insertion.

            // Populate the chart with custom data.
            Chart chart = chartShape.Chart;
            chart.Series.Clear();
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[] values = { 15.0, 30.0, 25.0, 40.0 };
            chart.Series.Add("Sales", categories, values);

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartInTable.docx");
            doc.Save(outputPath);
        }
    }
}
