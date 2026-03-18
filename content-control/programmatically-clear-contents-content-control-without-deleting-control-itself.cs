using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a plain‑text content control (SDT) into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        builder.InsertNode(sdt);

        // Add some text inside the content control.
        Run run = new Run(doc, "Initial content");
        sdt.AppendChild(run);

        // At this point the content control contains text.
        Console.WriteLine("Before clear: " + sdt.GetText().Trim());

        // Clear the contents of the content control while keeping the control itself.
        // This uses the StructuredDocumentTag.Clear method as defined in the API.
        sdt.Clear();

        // After clearing, the control remains and may show placeholder text if defined.
        Console.WriteLine("After clear: " + sdt.GetText().Trim());

        // Save the document to verify the result.
        doc.Save("ContentControlCleared.docx");
    }
}
