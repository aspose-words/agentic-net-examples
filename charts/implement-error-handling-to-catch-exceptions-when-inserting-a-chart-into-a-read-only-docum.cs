using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a new document and insert a simple chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape initialChart = builder.InsertChart(ChartType.Column, 400, 300);

        // Optional: clear demo data and add custom series.
        Chart chart = initialChart.Chart;
        chart.Series.Clear();
        chart.Series.Add("Sample", new[] { "A", "B", "C" }, new[] { 10.0, 20.0, 30.0 });

        // Step 2: Save the document to a writable memory stream.
        using (MemoryStream writableStream = new MemoryStream())
        {
            // Save to stream using the overload that specifies the format.
            doc.Save(writableStream, SaveFormat.Docx);
            // Ensure the stream's position is at the beginning for reading.
            writableStream.Position = 0;

            // Step 3: Create a read‑only stream from the same byte array.
            MemoryStream readOnlyStream = new MemoryStream(writableStream.ToArray(), writable: false);

            // Step 4: Load the document from the read‑only stream.
            Document readOnlyDoc = new Document(readOnlyStream);
            DocumentBuilder readOnlyBuilder = new DocumentBuilder(readOnlyDoc);

            // Step 5: Attempt to modify the document (insert another chart) and handle any exceptions.
            try
            {
                readOnlyBuilder.InsertChart(ChartType.Pie, 300, 300);
                Console.WriteLine("Chart inserted successfully into the read‑only document object.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception during chart insertion: {ex.Message}");
            }

            // Step 6: Attempt to save the modified document back to the same read‑only stream.
            try
            {
                // Reset position to the start; this will fail because the stream is not writable.
                readOnlyStream.Position = 0;
                readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);
                Console.WriteLine("Document saved back to the read‑only stream (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception during saving to read‑only stream: {ex.Message}");
            }

            // Step 7: Save the final document to a regular file to verify the result.
            readOnlyDoc.Save("Result.docx");
        }
    }
}
