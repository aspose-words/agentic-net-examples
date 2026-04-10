using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartInTableCellExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // Define the desired cell dimensions (points).
        // Width of the cell.
        builder.CellFormat.Width = 300; // 300 points ≈ 4.17 cm
        // Height of the row (which determines cell height).
        builder.RowFormat.Height = 200; // 200 points ≈ 2.78 cm
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Insert a column chart that fits exactly into the cell.
        // The InsertChart method returns a Shape that represents the chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 300, 200);
        Chart chart = chartShape.Chart;

        // Clear the default demo data.
        chart.Series.Clear();

        // Add custom data to the chart.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120, 150, 180, 130 };
        chart.Series.Add("Sales", categories, values);

        // Optionally set a title for the chart.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the working directory.
        doc.Save("ChartInTableCell.docx");
    }
}
