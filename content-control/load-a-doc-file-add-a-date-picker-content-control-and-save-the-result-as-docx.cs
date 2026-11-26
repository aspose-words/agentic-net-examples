using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOC file that will act as the existing source.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        Paragraph sourceParagraph = new Paragraph(sourceDoc);
        sourceParagraph.AppendChild(new Run(sourceDoc, "This is a sample DOC file."));
        sourceDoc.FirstSection.Body.AppendChild(sourceParagraph);
        sourceDoc.Save("input.doc", SaveFormat.Doc);

        // --------------------------------------------------------------
        // 2. Load the existing DOC file that we just created.
        // --------------------------------------------------------------
        Document doc = new Document("input.doc");

        // --------------------------------------------------------------
        // 3. Create an inline plain‑text content control that will hold a
        //    date value (Aspose.Words versions prior to supporting a
        //    dedicated DateTime SDT use a plain‑text SDT for this purpose).
        // --------------------------------------------------------------
        StructuredDocumentTag dateControl = new StructuredDocumentTag(
            doc,
            SdtType.PlainText,          // Use a plain‑text SDT as a fallback.
            MarkupLevel.Inline);

        dateControl.Title = "DatePicker";
        dateControl.Tag = "date-picker";

        // Remove any default children and insert the formatted date string.
        dateControl.RemoveAllChildren();
        string formattedDate = DateTime.Today.ToString("yyyy-MM-dd");
        dateControl.AppendChild(new Run(doc, formattedDate));

        // --------------------------------------------------------------
        // 4. Insert the content control into a new paragraph.
        // --------------------------------------------------------------
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "Selected date: "));
        paragraph.AppendChild(dateControl);
        doc.FirstSection.Body.AppendChild(paragraph);

        // --------------------------------------------------------------
        // 5. Save the modified document as DOCX.
        // --------------------------------------------------------------
        doc.Save("output.docx", SaveFormat.Docx);
    }
}
