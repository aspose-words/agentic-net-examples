using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string docxPath = "sample.docx";
        const string pdfPath = "sample.pdf";
        const string outputDocxPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a Word document with a combo box form field.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Please select a fruit: ");
        builder.InsertComboBox("MyComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Convert the Word document to PDF while preserving form fields.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };
        sourceDoc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF file was not created.");

        // -----------------------------------------------------------------
        // 3. Load the PDF and convert it back to DOCX.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(outputDocxPath, SaveFormat.Docx);

        // Verify that the output DOCX was created.
        if (!File.Exists(outputDocxPath))
            throw new InvalidOperationException("Expected output DOCX file was not created.");
    }
}
