using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the intermediate PDF and final DOCX.
        const string pdfPath = "sample_form.pdf";
        const string docxPath = "converted.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a Word document with a combo box form field.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Write("Please select a fruit: ");
        builder.InsertComboBox("MyComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);

        // Save the document as PDF, preserving the form fields as interactive PDF fields.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };
        sourceDoc.Save(pdfPath, pdfSaveOptions);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF that contains the form fields.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // Step 3: Convert the PDF to DOCX. Form fields become content controls.
        // -----------------------------------------------------------------
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: Ensure the DOCX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created.");

        // Optional: Clean up intermediate PDF if desired.
        // File.Delete(pdfPath);
    }
}
