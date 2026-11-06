using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF file.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF content created for conversion.");
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF file was not created.");

        // Step 2: Load the PDF file.
        Document pdfDocument = new Document(pdfPath);

        // Step 3: Convert and save as DOCX.
        string docxPath = "converted.docx";
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Verify the DOCX was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created after conversion.");

        // Optional: Clean up generated files (comment out if you want to keep them).
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
