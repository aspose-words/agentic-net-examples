using System;
using Aspose.Words;

class InsertParagraphsInLoop
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Loop to insert multiple paragraphs.
        for (int i = 1; i <= 5; i++)
        {
            // Write custom text and end the paragraph.
            builder.Writeln($"This is paragraph number {i}.");
        }

        // Save the document to a file.
        doc.Save("InsertedParagraphs.docx");
    }
}
