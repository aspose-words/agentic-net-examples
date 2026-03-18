using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace BatchChartInserter
{
    public class WordBatchChartInserter
    {
        // Path to the folder containing the Word documents.
        private readonly string _folderPath;

        // Constructor – creates an instance with the target folder.
        public WordBatchChartInserter(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("Folder path must be a valid non‑empty string.", nameof(folderPath));

            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException($"The folder \"{folderPath}\" does not exist.");

            _folderPath = folderPath;
        }

        // Executes the batch processing: adds a predefined bar chart to the first page of each .docx file.
        public void Execute()
        {
            // Get all Word documents in the folder (non‑recursive).
            string[] docFiles = Directory.GetFiles(_folderPath, "*.docx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in docFiles)
            {
                // Load the existing document.
                Document doc = new Document(filePath);

                // Use DocumentBuilder to insert content.
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Move the cursor to the very start of the document (first page).
                builder.MoveToDocumentStart();

                // Insert a bar chart with a fixed size (width: 400 points, height: 300 points).
                // The InsertChart method returns a Shape that contains the chart.
                Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

                // Access the Chart object to define its data.
                Chart chart = chartShape.Chart;

                // Clear any default series and add our predefined series.
                chart.Series.Clear();

                // Example data: three categories with corresponding values.
                string[] categories = { "Category A", "Category B", "Category C" };
                double[] values = { 10.0, 20.0, 30.0 };

                // Add a single series named "Series 1".
                chart.Series.Add("Series 1", categories, values);

                // Optional: set a title for the chart.
                chart.Title.Text = "Sample Bar Chart";
                chart.Title.Show = true;

                // Save the modified document, overwriting the original file.
                doc.Save(filePath);
            }
        }
    }

    public static class Program
    {
        // Entry point required for a console application.
        public static void Main(string[] args)
        {
            // Use the first command‑line argument as the folder path, or fall back to a hard‑coded example.
            string folderPath = args.Length > 0 ? args[0] : @"C:\Docs\WordFiles";

            var inserter = new WordBatchChartInserter(folderPath);
            inserter.Execute();
        }
    }
}
