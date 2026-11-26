using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First cell (optional label).
        builder.InsertCell();
        builder.Write("Chart:");

        // Second cell where the chart will be placed.
        builder.InsertCell();

        // Desired dimensions for the cell and the chart (in points).
        double cellWidth = 300;
        double cellHeight = 180;

        // Set the cell width so the chart fits proportionally.
        builder.CellFormat.Width = cellWidth;

        // Insert a column chart with the same dimensions as the cell.
        Shape chartShape = builder.InsertChart(ChartType.Column, cellWidth, cellHeight);
        Chart chart = chartShape.Chart;

        // Populate the chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Sales", new double[] { 15, 30, 45 });

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartInTable.docx");
        doc.Save(outputPath);
    }
}
