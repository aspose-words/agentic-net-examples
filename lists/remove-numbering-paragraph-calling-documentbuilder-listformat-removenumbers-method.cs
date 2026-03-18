using System;
using Aspose.Words;
using Aspose.Words.Lists;

class RemoveListNumbering
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a default numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Remove numbering from the third paragraph only.
        // Retrieve the third paragraph (index 2) from the document.
        Paragraph thirdParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[2];
        // This call clears the list formatting for that specific paragraph.
        thirdParagraph.ListFormat.RemoveNumbers();

        // End list formatting for any subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph after the list.
        builder.Writeln("This paragraph is not part of a list.");

        // Save the document to disk.
        doc.Save("RemoveNumbers.docx");
    }
}
