using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new document and insert a chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape initialChart = builder.InsertChart(ChartType.Column, 432, 252);

        // Save the document to a byte array – this will be the source for the read‑only stream.
        byte[] docBytes;
        using (MemoryStream ms = new MemoryStream())
        {
            doc.Save(ms, SaveFormat.Docx);
            docBytes = ms.ToArray();
        }

        // -----------------------------------------------------------------
        // 2. Load the document from a read‑only stream.
        //    MemoryStream(byte[], bool) with 'false' creates a non‑writable stream.
        // -----------------------------------------------------------------
        using (MemoryStream readOnlyStream = new MemoryStream(docBytes, false))
        {
            Document readOnlyDoc = new Document(readOnlyStream);
            DocumentBuilder roBuilder = new DocumentBuilder(readOnlyDoc);

            try
            {
                // Attempt to modify the document by inserting another chart.
                roBuilder.InsertChart(ChartType.Pie, 300, 300);

                // Trying to save back to the same read‑only stream will raise an exception.
                readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);

                // If no exception occurs (unexpected), inform the user.
                Console.WriteLine("Document saved to read‑only stream (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected path: the stream does not support writing.
                Console.WriteLine($"Caught exception: {ex.GetType().Name} - {ex.Message}");
            }

            // Save the modified document to a regular file to demonstrate that the
            // in‑memory changes are valid despite the read‑only source stream.
            readOnlyDoc.Save("ModifiedFromReadOnly.docx");
        }

        // -----------------------------------------------------------------
        // 3. Normal save of the original document (optional demonstration).
        // -----------------------------------------------------------------
        doc.Save("OriginalDocument.docx");
    }
}
