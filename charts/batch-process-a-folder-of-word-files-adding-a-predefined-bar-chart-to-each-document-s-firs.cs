using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;        // Chart related classes

public class Program
{
    public static void Main()
    {
        // Define a working directory for the batch operation.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ChartsBatch");
        Directory.CreateDirectory(workDir);

        // Create a few sample Word documents if the folder is empty.
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

            // Position the builder at the start of the document (first page).
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentStart();

            // Insert a bar chart onto the first page.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 432, 252);
            Chart chart = chartShape.Chart;

            // Remove the demo data that Aspose.Words inserts by default.
            chart.Series.Clear();

            // Define categories and values for the predefined chart.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[] values = { 15000, 23000, 18000, 27000 };

            // Add a single series with the data.
            chart.Series.Add("Quarterly Sales", categories, values);

            // Set a visible title for the chart.
            chart.Title.Text = "Quarterly Sales Overview";
            chart.Title.Show = true;

            // Save the modified document back to the same file.
            doc.Save(filePath);
        }
    }
}
