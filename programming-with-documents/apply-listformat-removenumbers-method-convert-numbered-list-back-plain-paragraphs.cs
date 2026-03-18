using Aspose.Words;
using Aspose.Words.Lists;
using System;
using System.Linq;

class RemoveListNumbers
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list and add three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // End the list so further paragraphs are not automatically numbered.
        builder.ListFormat.RemoveNumbers();

        // Add a regular paragraph after the list.
        builder.Writeln("This paragraph is not part of a list.");

        // Remove list formatting from all paragraphs that are currently list items.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true)
                                      .OfType<Paragraph>()
                                      .Where(p => p.ListFormat.IsListItem))
        {
            para.ListFormat.RemoveNumbers();
        }

        // Save the document.
        doc.Save("RemoveNumbers.docx");
    }
}
