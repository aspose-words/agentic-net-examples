using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartInTableExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First cell – label.
        builder.InsertCell();
        builder.Write("Chart:");

        // Define desired dimensions for the chart cell (in points).
        const double cellWidth = 300;   // ~4.17 inches
        const double cellHeight = 200;  // ~2.78 inches

        // Apply the dimensions to the upcoming cell and its row.
        builder.CellFormat.Width = cellWidth;
        builder.RowFormat.Height = cellHeight;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Second cell – will contain the chart.
        builder.InsertCell();

        // Insert a column chart that fits the cell size.
        Shape chartShape = builder.InsertChart(ChartType.Column, cellWidth, cellHeight);
        // Note: AspectRatioLocked caused a NullReference in some environments,
        // so it is omitted. The chart will already respect the specified size.

        // Populate the chart with sample data.
        Chart chart = chartShape.Chart;
        chart.Series.Clear();
        chart.Series.Add("Sales",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] { 150.0, 200.0, 180.0, 220.0 });

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("ChartInTable.docx");
    }
}
