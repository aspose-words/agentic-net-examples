using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a plain‑text content control that shows placeholder text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will contain the content control.
        builder.Writeln("Please fill the following field:");

        // Create an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name",
            IsShowingPlaceholderText = true // Show placeholder when the control is empty.
        };
        // Define placeholder text.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Enter name here"));

        // Append the content control to the current paragraph.
        builder.InsertNode(sdt);
        builder.Writeln(); // Move to next line.

        // Save the DOCX to the working directory.
        const string docxPath = "sample-with-content-controls.docx";
        doc.Save(docxPath);

        // Step 2: Load the DOCX and convert it to PDF while preserving the placeholder appearance.
        Document loadedDoc = new Document(docxPath);

        // Configure PDF save options if needed (e.g., preserve form fields).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = false // Content controls are not form fields; keep as static text.
        };

        // Save as PDF.
        const string pdfPath = "converted.pdf";
        loadedDoc.Save(pdfPath, pdfOptions);
    }
}
