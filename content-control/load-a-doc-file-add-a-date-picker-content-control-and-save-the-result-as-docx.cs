using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a simple DOC (legacy Word 97-2003) document as the source file.
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document.");
        sourceDoc.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        var doc = new Document("input.doc");

        // Create an inline content control that will hold a date string.
        // The older Aspose.Words versions do not expose a dedicated DateTime SDT type,
        // so we use a plain text (inline) content control and populate it with the current date.
        var dateSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleDate",
            Tag = "sample-date"
        };

        // Remove any default children and add a Run with the formatted date.
        dateSdt.RemoveAllChildren();
        dateSdt.AppendChild(new Run(doc, DateTime.Now.ToString("yyyy-MM-dd")));

        // Insert the content control into the first paragraph.
        var firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(dateSdt);

        // Save the result as DOCX.
        doc.Save("output.docx", SaveFormat.Docx);
    }
}
