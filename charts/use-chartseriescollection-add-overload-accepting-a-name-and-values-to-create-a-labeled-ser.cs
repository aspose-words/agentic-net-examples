using System;
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

        // Insert a column chart with specific dimensions.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Remove any default series that may exist.
        chart.Series.Clear();

        // Add a labeled series using the overload that accepts a name and an array of values.
        chart.Series.Add("Quarterly Sales", new double[] { 15000, 20000, 25000, 30000 });

        // Save the document to the working directory.
        doc.Save("LabeledSeriesChart.docx");
    }
}
