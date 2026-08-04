using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Needed for TextFormFieldType

public class Program
{
    public static void Main()
    {
        // Paths for the intermediate PDF and the final DOCX.
        const string pdfPath = "sample_form.pdf";
        const string docxPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a Word document with form fields.
        // -------------------------------------------------
        Document wordDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordDoc);

        // Add a plain paragraph.
        builder.Writeln("Please fill in the form below:");

        // Insert a text input form field.
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Default text", 50);

        // Insert a combo box form field.
        builder.InsertComboBox("ComboBox", new[] { "Option A", "Option B", "Option C" }, 0);

        // Insert a check box form field.
        builder.InsertCheckBox("CheckBox", true, 0);

        // -------------------------------------------------
        // 2. Save the document as PDF while preserving form fields.
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };
        wordDoc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // -------------------------------------------------
        // 3. Load the PDF and convert it to DOCX.
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created.");

        // -------------------------------------------------
        // 4. Clean up (optional). Comment out if you want to inspect the files.
        // -------------------------------------------------
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
