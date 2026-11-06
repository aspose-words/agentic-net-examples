using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell where the chart will be placed.
        builder.StartTable();
        builder.InsertCell();

        // Obtain the current cell to set its dimensions.
        Cell? cell = builder.CurrentParagraph?.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current node is not a table cell.");

        // Define the desired width of the cell (in points).
        const double cellWidth = 300.0;
        cell.CellFormat.Width = cellWidth;

        // Insert a column chart into the cell.
        Shape chartShape = builder.InsertChart(ChartType.Column, cellWidth, 180.0);

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("Inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Preserve the original aspect ratio of the chart and scale it to the cell width.
        double originalRatio = chartShape.Height / chartShape.Width;
        chartShape.Width = cellWidth;
        chartShape.Height = chartShape.Width * originalRatio;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("ChartInTable.docx");
    }
}
