using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the working directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
                sampleBuilder.Writeln($"This is sample document {i}.");
                string samplePath = Path.Combine(inputFolder, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to the start of the document (first page).
            builder.MoveToDocumentStart();

            // Insert a bar chart onto the first page.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);
            Chart chart = chartShape.Chart;

            // Remove the demo data series.
            chart.Series.Clear();

            // Define categories and values for the chart.
            string[] categories = { "Category A", "Category B", "Category C" };
            double[] values = { 10.0, 20.0, 30.0 };

            // Add a new series with the defined data.
            chart.Series.Add("Sample Series", categories, values);

            // Set a visible title for the chart.
            chart.Title.Text = "Sample Bar Chart";
            chart.Title.Show = true;

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }
}
