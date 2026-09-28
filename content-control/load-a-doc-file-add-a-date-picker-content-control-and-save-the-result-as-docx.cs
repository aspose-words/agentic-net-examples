using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC (Word 97‑2003) file.
        const string docPath = "sample.doc";
        Document seedDoc = new Document();
        seedDoc.FirstSection.Body.FirstParagraph.AppendChild(new Run(seedDoc, "This is a sample DOC file."));
        seedDoc.Save(docPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(docPath);

        // Create a block‑level rich‑text content control that will act as a date picker placeholder.
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "DatePicker",
            Tag = "date-picker"
        };

        // Add a placeholder paragraph inside the content control.
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "Select a date"));
        dateSdt.AppendChild(placeholder);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(dateSdt);

        // Save the result as DOCX.
        const string outputPath = "output.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
