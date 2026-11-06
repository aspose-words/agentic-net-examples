using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;   // Required for Chart, ChartType, etc.

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with the desired size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove any automatically added default series.
        chart.Series.Clear();

        // Add a labeled series in a single call using the overload that accepts a name and values.
        chart.Series.Add("Quarterly Sales", new double[] { 15000, 20000, 18000, 22000 });

        // Save the document.
        doc.Save("LabeledSeriesChart.docx");
    }
}
