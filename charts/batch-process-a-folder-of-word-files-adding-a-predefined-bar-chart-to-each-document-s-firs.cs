using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define a working directory for the batch operation.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ChartsBatch");
        Directory.CreateDirectory(workDir);

        // Create a few sample DOCX files if the folder is empty.
        // This ensures the example is self‑contained.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(workDir, $"input-{i}.docx");
            if (!File.Exists(samplePath))
            {
                Document sampleDoc = new Document();
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOCX file in the folder.
        foreach (string filePath in Directory.GetFiles(workDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Prepare a builder positioned at the start of the document (first page).
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentStart();

            // Insert a Bar chart with a predefined size.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 432, 252);
            Chart chart = chartShape.Chart;

            // Clear any demo data that comes with the inserted chart.
            chart.Series.Clear();

            // Define categories and values for the chart.
            string[] categories = { "Category 1", "Category 2", "Category 3" };
            double[] values = { 10, 20, 30 };

            // Add a single series to the chart.
            chart.Series.Add("Sample Series", categories, values);

            // Add a visible title to the chart.
            chart.Title.Text = "Sample Bar Chart";
            chart.Title.Show = true;
            chart.Title.Font.Size = 14;
            chart.Title.Font.Color = Color.Blue;

            // Save the modified document back to the same file.
            doc.Save(filePath);
        }
    }
}
