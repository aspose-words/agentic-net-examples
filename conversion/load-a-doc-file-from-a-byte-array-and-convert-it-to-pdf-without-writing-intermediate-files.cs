using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOC document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample DOC document created in memory.");

        // Step 2: Save the DOC document to a MemoryStream (no file is written).
        byte[] docBytes;
        using (MemoryStream docStream = new MemoryStream())
        {
            sourceDoc.Save(docStream, SaveFormat.Doc);
            docBytes = docStream.ToArray(); // Capture the byte array.
        }

        // Step 3: Load a new Document from the byte array.
        Document loadedDoc;
        using (MemoryStream loadStream = new MemoryStream(docBytes))
        {
            loadedDoc = new Document(loadStream);
        }

        // Step 4: Convert the loaded DOC document to PDF and save to a file.
        const string outputPdfPath = "output.pdf";
        loadedDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Step 5: Validate that the PDF file was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: Output a simple confirmation (no interactive wait).
        Console.WriteLine("PDF conversion completed successfully.");
    }
}
