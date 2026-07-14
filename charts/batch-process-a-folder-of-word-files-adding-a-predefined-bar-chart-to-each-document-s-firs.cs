using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart related types

public class Program
{
    public static void Main()
    {
        // Define a working folder for the batch operation.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ChartsBatch");
        Directory.CreateDirectory(workDir);

        // Create a few sample Word documents if the folder is empty.
        // Each document will contain a simple paragraph.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(workDir, $"input-{i}.docx");
            if (!File.Exists(samplePath))
            {
                Document sampleDoc = new Document();
                DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
                sampleBuilder.Writeln($"Sample document {i}");
                sampleDoc.Save(samplePath);
            }
        }

        // Process every DOCX file in the folder.
        foreach (string filePath in Directory.GetFiles(workDir, "*.docx"))
        {
            // Load the existing document.
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the very start of the document (first page).
            builder.MoveToDocumentStart();

            // Insert a bar chart and obtain its Chart object.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 432, 252);
            Chart chart = chartShape.Chart;

            // Configure the chart title.
            chart.Title.Text = "Quarterly Sales";
            chart.Title.Show = true;

            // Remove the demo data that comes with a new chart.
            chart.Series.Clear();

            // Define categories and corresponding values for the bar chart.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[] values = { 1500, 2000, 1800, 2200 };

            // Add a single series with the predefined data.
            chart.Series.Add("Sales", categories, values);

            // Save the modified document back to the same file.
            doc.Save(filePath);
        }
    }
}
