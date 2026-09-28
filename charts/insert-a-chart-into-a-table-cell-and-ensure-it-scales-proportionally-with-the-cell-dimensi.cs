using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

public class ChartInTableExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First cell – just some text.
        builder.InsertCell();
        builder.Write("Sales Data");

        // Second cell – will contain the chart.
        builder.InsertCell();

        // Define the original chart size (points).
        const double originalChartWidth = 300.0;
        const double originalChartHeight = 180.0;

        // Desired cell width (points). Height will be calculated to keep the aspect ratio.
        const double desiredCellWidth = 400.0;

        // Calculate scaling factor and the new chart height to keep the proportion.
        double scale = desiredCellWidth / originalChartWidth;
        double scaledChartHeight = originalChartHeight * scale;

        // Set the cell width for the current cell.
        Cell? chartCell = builder.CurrentParagraph?.ParentNode as Cell;
        if (chartCell != null)
        {
            chartCell.CellFormat.Width = desiredCellWidth;
        }

        // Insert the chart with the original size; we'll resize it after insertion.
        Shape chartShape = builder.InsertChart(ChartType.Column, originalChartWidth, originalChartHeight);
        Chart chart = chartShape.Chart;

        // Populate the chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Q1", new double[] { 120, 150, 180 });
        chart.Series.Add("Q2", new double[] { 130, 160, 190 });

        // Resize the chart to match the cell dimensions while preserving the aspect ratio.
        chartShape.Width = desiredCellWidth;
        chartShape.Height = scaledChartHeight;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartInTable.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
