using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Required for Shape
using Aspose.Words.Drawing.Charts;
using System.Drawing;

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
            // Ensure the document has at least one paragraph.
            sampleDoc.FirstSection.Body.FirstParagraph.AppendChild(new Run(sampleDoc, $"Sample document {i}"));
            sampleDoc.Save(Path.Combine(workDir, $"input-{i}.docx"));
        }

        // Process each DOCX file in the folder.
        foreach (string filePath in Directory.GetFiles(workDir, "*.docx"))
        {
            // Load the existing document.
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to the start of the document (first page).
            builder.MoveToDocumentStart();

            // Insert a bar chart.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 432, 252);
            Chart chart = chartShape.Chart;

            // Remove the demo data that comes with a new chart.
            chart.Series.Clear();

            // Define categories and values for the predefined chart.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[] values = { 10.0, 20.0, 30.0, 40.0 };

            // Add a single series with the custom data.
            chart.Series.Add("Sales", categories, values);

            // Configure the chart title.
            chart.Title.Text = "Quarterly Sales";
            chart.Title.Show = true;
            chart.Title.Font.Size = 14;
            chart.Title.Font.Color = Color.Blue;

            // Save the modified document back to the same file.
            doc.Save(filePath);
        }
    }
}
