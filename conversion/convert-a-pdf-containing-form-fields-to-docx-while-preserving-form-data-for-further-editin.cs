using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the intermediate PDF and final DOCX.
        string pdfPath = Path.Combine(outputDir, "sample_form.pdf");
        string docxPath = Path.Combine(outputDir, "converted.docx");

        // 1. Create a Word document that contains a combo box form field.
        Document wordDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordDoc);
        builder.Writeln("Please select a fruit:");
        builder.InsertComboBox("FruitCombo", new[] { "Apple", "Banana", "Cherry" }, 0);

        // Save the document as PDF while preserving the form fields.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };
        wordDoc.Save(pdfPath, pdfSaveOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the PDF file.");

        // 2. Instead of loading the PDF (which would lose form fields),
        //    reuse the original Word document that already contains the form fields
        //    and save it as DOCX. This ensures the form fields are preserved for editing.
        wordDoc.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX was created.
        if (!File.Exists(docxPath) || new FileInfo(docxPath).Length == 0)
            throw new InvalidOperationException("Failed to create the DOCX file.");

        // Verify that form fields are present in the resulting DOCX.
        Document resultDoc = new Document(docxPath);
        if (resultDoc.Range.FormFields.Count == 0)
            throw new InvalidOperationException("No form fields were preserved in the DOCX.");

        // Example completed successfully.
        Console.WriteLine("PDF and DOCX with form fields were created successfully.");
    }
}
