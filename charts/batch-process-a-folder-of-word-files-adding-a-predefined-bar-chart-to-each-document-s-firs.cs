using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a working folder for the batch operation.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ChartsBatch");
        Directory.CreateDirectory(workDir);

        // Generate a few sample Word documents to process.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            sampleDoc.Save(Path.Combine(workDir, $"input-{i}.docx"));
        }

        // Process each DOCX file in the folder, adding a predefined bar chart to the first page.
        foreach (string filePath in Directory.GetFiles(workDir, "*.docx"))
        {
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to the start of the document (first page) and insert a bar chart.
            builder.MoveToDocumentStart();
            Shape chartShape = builder.InsertChart(ChartType.Bar, 432, 252);
            Chart chart = chartShape.Chart;

            // Define custom data for the chart.
            chart.Series.Clear();
            chart.Series.Add("Sales", new double[] { 120, 150, 180 });

            // Save the modified document back to the same file.
            doc.Save(filePath);
        }
    }
}
