using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class StructuredDocumentTagExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph to host the structured document tag.
        builder.Writeln("Below is a checkbox content control:");

        // Insert a checkbox StructuredDocumentTag (content control) at the current cursor position.
        // The markup level is detected automatically (inline in this case).
        StructuredDocumentTag checkBox = builder.InsertStructuredDocumentTag(SdtType.Checkbox);

        // Set the initial state of the checkbox.
        checkBox.Checked = true;

        // Optionally set custom symbols for checked/unchecked states.
        checkBox.SetCheckedSymbol(0x00A9, "Times New Roman");   // ©
        checkBox.SetUncheckedSymbol(0x00AE, "Times New Roman"); // ®

        // Add some explanatory text after the checkbox.
        builder.Writeln();
        builder.Writeln("The document contains a structured document tag (SDT).");

        // Save the document as PDF with document structure exported.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true // Enables tags in the PDF for accessibility.
        };

        // Save the PDF to disk.
        doc.Save("StructuredDocumentTagExample.pdf", pdfOptions);
    }
}
