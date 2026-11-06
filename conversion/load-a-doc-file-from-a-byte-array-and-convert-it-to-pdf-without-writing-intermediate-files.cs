using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content.");

        // Save the DOC document to a memory stream (no file is written).
        using (MemoryStream docStream = new MemoryStream())
        {
            sourceDoc.Save(docStream, SaveFormat.Doc);
            // Ensure the stream position is at the beginning for reading.
            docStream.Position = 0;

            // Load a Document from the byte array contained in the memory stream.
            byte[] docBytes = docStream.ToArray();
            using (MemoryStream loadStream = new MemoryStream(docBytes))
            {
                Document loadedDoc = new Document(loadStream);

                // Convert the loaded DOC document to PDF and save to a file.
                const string pdfPath = "output.pdf";
                loadedDoc.Save(pdfPath, SaveFormat.Pdf);

                // Validate that the PDF file was created.
                if (!File.Exists(pdfPath))
                {
                    throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }
    }
}
