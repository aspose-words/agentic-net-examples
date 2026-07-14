using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string pdfPath = "sample_form.pdf";
        const string docxPath = "converted.docx";

        // -----------------------------------------------------------------
        // 1. Create a Word document with a form field (combo box) and save it as PDF.
        // -----------------------------------------------------------------
        Document wordDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordDoc);

        builder.Writeln("Please select a fruit:");
        // Insert a combo box form field with three options.
        builder.InsertComboBox("FruitCombo", new[] { "Apple", "Banana", "Cherry" }, 0);

        // Save as PDF while preserving the form fields as interactive PDF fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };
        wordDoc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to DOCX.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX file was not created.");
    }
}
