using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample document that will be saved as PDF.
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Sample content for PDF conversion.");

        // Save the document to a memory stream as PDF (simulating a PDF downloaded from a URL).
        using MemoryStream pdfStream = new MemoryStream();
        sourceDocument.Save(pdfStream, SaveFormat.Pdf);
        pdfStream.Position = 0; // Reset stream position for reading.

        // Load the PDF from the memory stream.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Pdf };
        Document pdfDocument = new Document(pdfStream, loadOptions);

        // Convert the loaded PDF to DOCX and save to a file.
        string outputPath = "output.docx";
        pdfDocument.Save(outputPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }
    }
}
