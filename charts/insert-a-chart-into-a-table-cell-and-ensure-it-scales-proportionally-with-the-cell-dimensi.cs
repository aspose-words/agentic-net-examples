using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

public class ChartInTableExample
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First cell – just some placeholder text.
        builder.InsertCell();
        builder.Write("Data cell");

        // Second cell – will contain the chart.
        builder.InsertCell();

        // Define explicit cell dimensions.
        // Width of the cell (points). 1 point = 1/72 inch.
        builder.CellFormat.Width = 300; // approx 4.17 cm
        // Height of the row (and thus the cell) – set to exact value.
        builder.RowFormat.HeightRule = HeightRule.Exactly;
        builder.RowFormat.Height = 200; // approx 2.78 cm

        // Insert a chart that matches the cell size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 300, 200);
        Chart chart = chartShape.Chart;

        // Optional: replace the demo data with custom series.
        chart.Series.Clear();
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 150, 200, 180, 220 });

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("ChartInTable.docx");
    }
}
