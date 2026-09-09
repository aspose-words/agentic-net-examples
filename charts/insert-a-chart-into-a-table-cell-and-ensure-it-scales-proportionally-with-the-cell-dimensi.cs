using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define cell dimensions (points).
        const double cellWidth = 300.0;
        const double cellHeight = 200.0;

        // Start a table.
        builder.StartTable();

        // First cell – just some placeholder text.
        builder.InsertCell();
        builder.CellFormat.Width = cellWidth;
        builder.RowFormat.Height = cellHeight;
        builder.RowFormat.HeightRule = HeightRule.Exactly;
        builder.Write("Placeholder");

        // End the first row.
        builder.EndRow();

        // Second cell – the chart will be inserted here.
        builder.InsertCell();
        builder.CellFormat.Width = cellWidth;
        builder.RowFormat.Height = cellHeight;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Insert a column chart that matches the cell size.
        Shape chartShape = builder.InsertChart(ChartType.Column, cellWidth, cellHeight);
        Chart chart = chartShape.Chart;

        // Remove the demo data and add custom series.
        chart.Series.Clear();
        chart.Series.Add(
            "Quarterly Sales",
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] { 120.0, 150.0, 180.0, 200.0 });

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("ChartInTable.docx");
    }
}
