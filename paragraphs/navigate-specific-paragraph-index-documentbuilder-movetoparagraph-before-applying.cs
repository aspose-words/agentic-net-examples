using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the start of the third paragraph (index = 2).
        // The second argument (characterIndex) is set to 0 to position at the beginning.
        builder.MoveToParagraph(2, 0);

        // Apply desired paragraph formatting at the new cursor location.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.ParagraphFormat.SpaceAfter = 12; // 12 points spacing after the paragraph.

        // Insert new text at the moved cursor position.
        builder.Writeln("This text is inserted into paragraph 3.");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
