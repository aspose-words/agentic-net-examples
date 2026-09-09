using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Column chart.
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart columnChart = columnChartShape.Chart;

        // Clear the demo data.
        columnChart.Series.Clear();

        // Define categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 150.0, 130.75, 170.25 };

        // Add a series with the data.
        columnChart.Series.Add("Sales", categories, values);

        // ----- Dynamic transformation: replace the column chart with a line chart -----
        // Remember the data for reuse.
        string[] savedCategories = (string[])categories.Clone();
        double[] savedValues = (double[])values.Clone();

        // Move the builder back to the column chart shape.
        builder.MoveTo(columnChartShape);
        // Remove the column chart shape from the document.
        columnChartShape.Remove();

        // Insert a Line chart at the same position.
        Shape lineChartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart lineChart = lineChartShape.Chart;

        // Clear any demo data and add the saved series.
        lineChart.Series.Clear();
        lineChart.Series.Add("Sales", savedCategories, savedValues);
        // ------------------------------------------------------------------------------

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicChart.docx");
        doc.Save(outputPath);
    }
}
