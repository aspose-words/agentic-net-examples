using System;
using Aspose.Words;
using Aspose.Words.Saving;

class CreatePdfWithFillableForms
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please select a fruit:");

        // Insert a combo box (drop‑down) form field.
        // The field will be preserved as an interactive PDF form field.
        builder.InsertComboBox("FruitComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);

        // Configure PDF save options to keep Word form fields as PDF form fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true,               // Export form fields as interactive objects.
            RenderChoiceFormFieldBorder = true       // Optional: render borders for choice fields.
        };

        // Save the document as a PDF with fillable forms.
        doc.Save("FillableForm.pdf", pdfOptions);
    }
}
