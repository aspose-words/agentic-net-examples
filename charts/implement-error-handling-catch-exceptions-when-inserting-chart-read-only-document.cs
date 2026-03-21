using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class InsertChartIntoReadOnlyStream
{
    static void Main()
    {
        // Use files relative to the program's working directory so the example works everywhere.
        string baseDir = AppContext.BaseDirectory;
        string sourcePath = Path.Combine(baseDir, "SourceDocument.docx");
        string outputPath = Path.Combine(baseDir, "ModifiedDocument.docx");

        // Ensure a source document exists. If it does not, create a minimal one.
        if (!File.Exists(sourcePath))
        {
            var emptyDoc = new Document();
            emptyDoc.Save(sourcePath);
        }

        // Open the source document as a read‑only stream.
        using (FileStream readOnlyStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            // Load the document from the read‑only stream.
            Document doc = new Document(readOnlyStream);
            DocumentBuilder builder = new DocumentBuilder(doc);

            try
            {
                // Insert a chart. This works even though the source stream is read‑only,
                // because the document is loaded into memory.
                Shape chartShape = builder.InsertChart(ChartType.Pie, 300, 300);

                // Customize the chart (add a sample series).
                Chart chart = chartShape.Chart;
                chart.Series.Clear();
                chart.Series.Add("Sample", new[] { "A", "B", "C" }, new[] { 10.0, 20.0, 30.0 });

                // Save the modified document.
                doc.Save(outputPath);
                Console.WriteLine("Chart inserted and document saved successfully.");
            }
            catch (Exception ex)
            {
                // Handle any errors that might occur during chart insertion.
                Console.WriteLine($"Failed to insert chart: {ex.GetType().Name} - {ex.Message}");
            }
        }
    }
}
