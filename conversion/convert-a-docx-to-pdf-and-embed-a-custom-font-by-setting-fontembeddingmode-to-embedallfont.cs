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
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample document with a custom font.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Configure PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF with the specified options.
        doc.Save("output.pdf", pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Verify that the PDF file is not empty.
        FileInfo pdfInfo = new FileInfo("output.pdf");
        if (pdfInfo.Length == 0)
        {
            throw new InvalidOperationException("The generated PDF file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("DOCX successfully converted to PDF with all fonts embedded.");
    }
}
