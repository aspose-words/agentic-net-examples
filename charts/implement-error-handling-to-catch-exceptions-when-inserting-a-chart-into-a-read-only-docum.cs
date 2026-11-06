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
        // Create a new document and insert an initial column chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertChart(ChartType.Column, 400, 300);

        // Save the document to a memory stream.
        using (MemoryStream tempStream = new MemoryStream())
        {
            doc.Save(tempStream, SaveFormat.Docx);
            byte[] docBytes = tempStream.ToArray();

            // Create a read‑only stream from the saved bytes.
            using (MemoryStream readOnlyStream = new MemoryStream(docBytes, writable: false))
            {
                // Load the document from the read‑only stream.
                Document readOnlyDoc = new Document(readOnlyStream);
                DocumentBuilder readOnlyBuilder = new DocumentBuilder(readOnlyDoc);

                try
                {
                    // Attempt to insert another chart (pie chart) into the read‑only document.
                    readOnlyBuilder.InsertChart(ChartType.Pie, 400, 300);

                    // Attempt to save changes back to the read‑only stream.
                    // This will throw because the stream does not support writing.
                    readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);
                }
                catch (Exception ex)
                {
                    // Display the exception that occurs due to the read‑only stream.
                    Console.WriteLine($"Exception caught while saving to a read‑only stream: {ex.Message}");
                }

                // Save the modified document to a regular file to demonstrate a successful save.
                readOnlyDoc.Save("modified.docx");
            }
        }
    }
}
