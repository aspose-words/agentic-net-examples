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
        // 1. Create a Word document with a form field and save it as PDF.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);

        // Add a simple paragraph.
        builder.Writeln("Please select a fruit:");

        // Insert a combo box form field with three options.
        builder.InsertComboBox("FruitComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);

        // Save as PDF while preserving the form fields.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };
        pdfSource.Save(pdfPath, pdfSaveOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // ---------------------------------------------------------------
        // 2. Load the PDF and convert it to DOCX, keeping the form data.
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created.");
    }
}
