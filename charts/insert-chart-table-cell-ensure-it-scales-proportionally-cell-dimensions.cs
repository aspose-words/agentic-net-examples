using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and set a fixed width for the first cell.
        Table table = builder.StartTable();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(250);
        builder.InsertCell();

        // Insert a chart into the current cell.
        // Width and height of 0 request 100 % scaling relative to the shape size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 0, 0);

        // Ensure the chart behaves as an inline object inside the cell.
        chartShape.IsLayoutInCell = true;

        // Resize the chart to fill the cell while preserving its aspect ratio.
        double targetWidth = builder.CellFormat.PreferredWidth.Value;
        chartShape.Width = targetWidth;
        chartShape.Height = targetWidth; // square chart – adjust as needed for other ratios.

        // Populate the chart with sample data.
        Chart chart = chartShape.Chart;
        chart.Series.Clear();
        chart.Series.Add("Series 1", new[] { "A", "B", "C" }, new[] { 1.0, 2.0, 3.0 });

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Prevent the table from auto‑adjusting its width so the cell keeps the exact size.
        table.AllowAutoFit = false;

        // Save the resulting document.
        doc.Save("ChartInTableCell.docx");
    }
}
