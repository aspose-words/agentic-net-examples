using System;
using System.IO;
using System.Linq;
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

        // Create sample input documents if the folder is empty.
        if (!Directory.GetFiles(workDir, "*.docx").Any())
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
                sampleBuilder.Writeln($"Sample document {i}");
                string samplePath = Path.Combine(workDir, $"input-{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOCX file in the folder.
        foreach (string filePath in Directory.GetFiles(workDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move to the start of the document (first page) and insert a bar chart.
            builder.MoveToDocumentStart();
            Shape chartShape = builder.InsertChart(ChartType.Bar, 432, 252);

            // Ensure the inserted shape actually contains a chart before modifying it.
            if (chartShape.HasChart)
            {
                Chart chart = chartShape.Chart;

                // Clear any default series.
                chart.Series.Clear();

                // Add a single series with predefined categories and values.
                chart.Series.Add(
                    "Sales",
                    new string[] { "Q1", "Q2", "Q3" },
                    new double[] { 15, 25, 35 });
            }

            // Save the modified document back to the same file.
            doc.Save(filePath);
        }
    }
}
