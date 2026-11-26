using System;
using System.IO;
using Aspose.Words;

public class PdfToDocxConversion
{
    public static void Main()
    {
        // Create a sample PDF file to act as the input.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF document.");
        builder.Writeln("It contains multiple lines of text to demonstrate layout preservation.");
        string pdfInputPath = "input.pdf";
        sampleDoc.Save(pdfInputPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfInputPath))
            throw new InvalidOperationException("The input PDF file was not created.");

        // Load the PDF document. Layout preservation is handled by default when loading PDFs.
        Document pdfDocument = new Document(pdfInputPath);

        // Convert and save the document as DOCX.
        string docxOutputPath = "output.docx";
        pdfDocument.Save(docxOutputPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxOutputPath))
            throw new InvalidOperationException("The output DOCX file was not created.");

        // Optionally, clean up the sample input file.
        File.Delete(pdfInputPath);
    }
}
