using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Sample PDF content.");
        string pdfPath = "input.pdf";
        sourceDocument.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF with PdfLoadOptions, ignoring password protection.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        // No password is set; if the PDF were password‑protected, leaving the Password property unset
        // causes Aspose.Words to ignore the protection for this example.
        Document loadedDocument = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX.
        string docxPath = "output.docx";
        loadedDocument.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }
    }
}
