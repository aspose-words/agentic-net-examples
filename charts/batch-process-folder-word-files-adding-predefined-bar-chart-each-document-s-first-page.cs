using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace BatchChartInserter
{
    public class WordBatchChartInserter
    {
        private readonly string _folderPath;

        public WordBatchChartInserter(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("Folder path must be a valid non‑empty string.", nameof(folderPath));

            // Ensure the folder exists; create it if necessary.
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            _folderPath = folderPath;
        }

        public void Execute()
        {
            string[] docFiles = Directory.GetFiles(_folderPath, "*.docx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in docFiles)
            {
                Document doc = new Document(filePath);
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToDocumentStart();

                Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);
                Chart chart = chartShape.Chart;
                chart.Series.Clear();

                string[] categories = { "Category A", "Category B", "Category C" };
                double[] values = { 10.0, 20.0, 30.0 };
                chart.Series.Add("Series 1", categories, values);

                chart.Title.Text = "Sample Bar Chart";
                chart.Title.Show = true;

                doc.Save(filePath);
            }
        }
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            string folderPath = args.Length > 0
                ? args[0]
                : Path.Combine(AppContext.BaseDirectory, "WordFiles");

            // Ensure the folder exists.
            Directory.CreateDirectory(folderPath);

            // If the folder is empty, create a simple sample document to process.
            if (Directory.GetFiles(folderPath, "*.docx").Length == 0)
            {
                var sampleDoc = new Document();
                var builder = new DocumentBuilder(sampleDoc);
                builder.Writeln("This is a sample document.");
                string samplePath = Path.Combine(folderPath, "Sample.docx");
                sampleDoc.Save(samplePath);
            }

            var inserter = new WordBatchChartInserter(folderPath);
            inserter.Execute();

            Console.WriteLine($"Processing completed. Documents are located in: {folderPath}");
        }
    }
}
