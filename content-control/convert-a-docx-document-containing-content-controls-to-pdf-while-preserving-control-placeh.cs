using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample DOCX and the resulting PDF.
        string docPath = "sample.docx";
        string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a DOCX document that contains a plain‑text content
        // control with a placeholder.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading paragraph.
        builder.Writeln("Document with a content control placeholder:");

        // Create an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
            // PlaceholderName is not set because it requires a building block in the glossary.
        };

        // Add placeholder text as the child run of the content control.
        sdt.RemoveAllChildren(); // Ensure the SDT is empty before adding content.
        sdt.AppendChild(new Run(doc, "Enter name"));

        // Append the content control to the current paragraph.
        builder.CurrentParagraph?.AppendChild(sdt);

        // Save the DOCX file.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX document and convert it to PDF.
        // The placeholder text is part of the content control's child nodes,
        // so it will appear in the PDF output.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Save as PDF using default options.
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}
