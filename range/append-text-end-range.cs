using System;
using Aspose.Words;
using Aspose.Words.Markup;

class AppendTextToRange
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text.
        builder.Writeln("Original text.");

        // Insert a plain‑text structured document tag (range) at the current cursor position.
        // The range start node will hold the content of the tag.
        StructuredDocumentTagRangeStart rangeStart = new StructuredDocumentTagRangeStart(doc, SdtType.PlainText);
        StructuredDocumentTagRangeEnd rangeEnd = new StructuredDocumentTagRangeEnd(doc, rangeStart.Id);

        // Place the range start before the first paragraph and the range end after it.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        doc.FirstSection.Body.InsertBefore(rangeStart, firstParagraph);
        doc.FirstSection.Body.InsertAfter(rangeEnd, firstParagraph);

        // Create a Run node that contains the text we want to append.
        Run appendedRun = new Run(doc, " Appended text.");

        // Append the Run to the end of the stdContent range of the StructuredDocumentTagRangeStart.
        // AppendChild adds the node after any existing children inside the range.
        rangeStart.AppendChild(appendedRun);

        // Optionally, verify the resulting document text.
        Console.WriteLine("Document text after appending:");
        Console.WriteLine(doc.GetText().Trim());

        // Save the document.
        doc.Save("AppendTextToRange.docx");
    }
}
