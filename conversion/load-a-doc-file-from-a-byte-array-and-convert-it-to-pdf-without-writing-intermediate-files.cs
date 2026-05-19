using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");

        // Save the DOC document to a memory stream (no file is written).
        using (MemoryStream docStream = new MemoryStream())
        {
            source.Save(docStream, SaveFormat.Doc);

            // Obtain the byte array representing the DOC file.
            byte[] docBytes = docStream.ToArray();

            // Load a new Document from the byte array.
            using (MemoryStream inputStream = new MemoryStream(docBytes))
            {
                Document doc = new Document(inputStream);

                // Convert the document to PDF and save to disk.
                const string pdfPath = "output.pdf";
                doc.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException("Expected output PDF was not created.");
            }
        }
    }
}
