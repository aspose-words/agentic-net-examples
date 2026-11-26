using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // 1. Create a simple document and save it to a writable memory stream.
        Document initialDoc = new Document();
        using (MemoryStream writableStream = new MemoryStream())
        {
            initialDoc.Save(writableStream, SaveFormat.Docx);

            // 2. Obtain the byte array of the document.
            byte[] docBytes = writableStream.ToArray();

            // 3. Create a read‑only memory stream from the byte array.
            using (MemoryStream readOnlyStream = new MemoryStream(docBytes, writable: false))
            {
                // 4. Load the document from the read‑only stream.
                Document doc = new Document(readOnlyStream);

                try
                {
                    // 5. Insert a chart into the document.
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
                    Chart chart = chartShape.Chart;

                    // (Optional) Add a simple series to the chart.
                    chart.Series.Clear();
                    chart.Series.Add("Sample", new double[] { 10, 20, 30 });

                    // 6. Attempt to save the modified document back to the same read‑only stream.
                    // This will throw because the stream does not support writing.
                    doc.Save(readOnlyStream, SaveFormat.Docx);
                }
                catch (Exception ex)
                {
                    // 7. Handle the exception that occurs due to the read‑only stream.
                    Console.WriteLine("An error occurred while inserting a chart into a read‑only document stream:");
                    Console.WriteLine($"{ex.GetType().Name}: {ex.Message}");
                }

                // 8. Save the successfully modified document to a regular file for verification (optional).
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
                doc.Save(outputPath);
                Console.WriteLine($"Document saved to: {outputPath}");
            }
        }
    }
}
