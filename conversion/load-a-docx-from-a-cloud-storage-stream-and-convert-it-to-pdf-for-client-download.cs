using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello from Aspose.Words!");

        // Simulate a cloud storage stream by saving the DOCX into a MemoryStream.
        using MemoryStream cloudStream = new MemoryStream();
        sourceDoc.Save(cloudStream, SaveFormat.Docx);
        cloudStream.Position = 0; // Reset for reading.

        // Load the DOCX from the simulated cloud stream.
        Document loadedDoc = new Document(cloudStream);

        // Convert the document to PDF and write it to a response stream.
        using MemoryStream responseStream = new MemoryStream();
        loadedDoc.Save(responseStream, SaveFormat.Pdf);

        // Validate that the PDF data was written.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("PDF conversion produced an empty stream.");

        // Optionally save the PDF to a local file for verification.
        const string outputPath = "output.pdf";
        File.WriteAllBytes(outputPath, responseStream.ToArray());

        if (!File.Exists(outputPath))
            throw new InvalidOperationException("PDF file was not created.");
    }
}
