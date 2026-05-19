using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart types, styles, and related classes

public class ApplyChartStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart and apply the predefined "Blue" style.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300, ChartStyle.Blue);
        Chart chart = chartShape.Chart;

        // Replace the demo data with custom series.
        chart.Series.Clear();
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Sales", categories, new double[] { 15000, 20000, 18000, 22000 });

        // Save the document to the working directory.
        doc.Save("styled-chart.docx");
    }
}
