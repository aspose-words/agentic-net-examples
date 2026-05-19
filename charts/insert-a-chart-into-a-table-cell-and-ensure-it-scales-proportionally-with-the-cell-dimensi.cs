using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the desired cell dimensions (points).
        const double cellWidth = 300;   // Width of the cell and chart.
        const double cellHeight = 200;  // Height of the row (and chart).

        // Start a table.
        builder.StartTable();

        // First cell – just some placeholder text.
        builder.CellFormat.Width = cellWidth;
        builder.RowFormat.Height = cellHeight;
        builder.RowFormat.HeightRule = HeightRule.Exactly;
        builder.InsertCell();
        builder.Write("Chart will appear in the next cell.");

        // Second cell – insert the chart.
        builder.InsertCell();

        // Insert a chart with the same dimensions as the cell to keep proportional scaling.
        Shape chartShape = builder.InsertChart(ChartType.Column, cellWidth, cellHeight);
        Chart chart = chartShape.Chart;

        // Remove the demo data and add custom series.
        chart.Series.Clear();
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 15, 30, 45, 20 };
        chart.Series.Add("Quarterly Sales", categories, values);

        // Optional: set a title for clarity.
        chart.Title.Text = "Sales Overview";
        chart.Title.Show = true;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("ChartInTable.docx");
    }
}
