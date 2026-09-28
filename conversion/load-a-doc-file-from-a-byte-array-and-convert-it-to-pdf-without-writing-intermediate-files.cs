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
        builder.Writeln("Hello from in‑memory DOC.");

        // Save the source document to a MemoryStream in DOC format.
        using MemoryStream docStream = new MemoryStream();
        sourceDoc.Save(docStream, SaveFormat.Doc);
        byte[] docBytes = docStream.ToArray();

        // Load a new Document from the byte array.
        using MemoryStream loadStream = new MemoryStream(docBytes);
        Document loadedDoc = new Document(loadStream);

        // Convert the loaded document to PDF and save to a file.
        string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created and is not empty.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        FileInfo info = new FileInfo(pdfPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The generated PDF file is empty.");
        }
    }
}
