using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory("output");

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add initial text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Before control:");

        // Create a block‑level plain‑text content control (StructuredDocumentTag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);

        // Add a paragraph with some initial text inside the content control.
        Paragraph innerParagraph = new Paragraph(doc);
        innerParagraph.AppendChild(new Run(doc, "Initial content"));
        sdt.AppendChild(innerParagraph);

        // Insert the content control into the document body (valid location for block‑level SDT).
        doc.FirstSection.Body.AppendChild(sdt);

        // Move the builder after the inserted content control and add a line break.
        builder.MoveTo(sdt);
        builder.Writeln();

        // Save the document before clearing the control's contents.
        doc.Save("output/BeforeClear.docx");

        // Clear the contents of the content control while keeping the control itself.
        sdt.Clear();

        // Save the document after clearing the control's contents.
        doc.Save("output/AfterClear.docx");
    }
}
