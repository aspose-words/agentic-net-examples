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

        // Insert a column chart. The chart comes with three demo series by default.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Determine the index of the series to remove.
        // For this example we will remove the second series (zero‑based index 1) if it exists.
        int indexToRemove = 1;
        if (indexToRemove >= 0 && indexToRemove < chart.Series.Count)
        {
            chart.Series.RemoveAt(indexToRemove);
        }
        else
        {
            throw new InvalidOperationException($"Series index {indexToRemove} is out of range.");
        }

        // Save the document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RemoveSeries.docx");
        doc.Save(outputPath);
    }
}
