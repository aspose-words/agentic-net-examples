using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host inline content controls.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Plain‑text content control with placeholder text.
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainTextSdt.Title = "CustomerName";
        plainTextSdt.Tag = "customer-name";
        // Add placeholder text directly inside the SDT.
        plainTextSdt.AppendChild(new Run(doc, "Enter customer name"));
        // Insert the inline SDT into the paragraph.
        paragraph.AppendChild(plainTextSdt);

        // Rich‑text (block‑level) content control with placeholder text.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richTextSdt.Title = "Comments";
        richTextSdt.Tag = "comments";

        // Add a paragraph inside the block‑level content control containing placeholder text.
        Paragraph innerParagraph = new Paragraph(doc);
        innerParagraph.AppendChild(new Run(doc, "Enter comments here"));
        richTextSdt.AppendChild(innerParagraph);

        // Append the block‑level content control to the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the source DOCX file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the DOCX document that contains the content controls.
        Document loadedDoc = new Document(inputPath);

        // Convert the document to PDF while preserving placeholders.
        const string outputPath = "output.pdf";
        loadedDoc.Save(outputPath, SaveFormat.Pdf);
    }
}
