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
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with chart insertion test.");

        // Save the document to a writable memory stream.
        using (MemoryStream writableStream = new MemoryStream())
        {
            doc.Save(writableStream, SaveFormat.Docx);
            byte[] docBytes = writableStream.ToArray();

            // Create a read‑only stream from the byte array.
            using (MemoryStream readOnlyStream = new MemoryStream(docBytes, writable: false))
            {
                // Load the document from the read‑only stream.
                Document readOnlyDoc = new Document(readOnlyStream);

                try
                {
                    // Attempt to insert a chart.
                    DocumentBuilder roBuilder = new DocumentBuilder(readOnlyDoc);
                    Shape chartShape = roBuilder.InsertChart(ChartType.Column, 432, 252);
                    Chart chart = chartShape.Chart;
                    chart.Series.Add("Series 1", new double[] { 10, 20, 30 });

                    // Attempt to save back to the same read‑only stream (will throw).
                    readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);
                    Console.WriteLine("Chart inserted and document saved successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception caught while inserting chart into read‑only document: {ex.Message}");
                }

                // Save the document to a regular file to demonstrate successful output.
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
                readOnlyDoc.Save(outputPath);
                Console.WriteLine($"Document saved to {outputPath}");
            }
        }
    }
}
