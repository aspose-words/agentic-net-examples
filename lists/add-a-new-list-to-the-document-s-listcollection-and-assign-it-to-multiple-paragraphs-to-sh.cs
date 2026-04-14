using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 1 – normal text.");
        builder.Writeln("Paragraph 2 – will become a list item.");
        builder.Writeln("Paragraph 3 – will become a list item.");
        builder.Writeln("Paragraph 4 – normal text.");

        // Create a new list (bulleted) and add it to the document's ListCollection.
        List sharedList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Retrieve all paragraphs in the document.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .OfType<Paragraph>()
                           .ToList();

        // Assign the same list to Paragraph 2 and Paragraph 3.
        // This makes both paragraphs share the same list formatting.
        for (int i = 1; i <= 2; i++) // indices 1 and 2 correspond to Paragraph 2 and 3
        {
            paragraphs[i].ListFormat.List = sharedList;
            paragraphs[i].ListFormat.ListLevelNumber = 0; // first level of the list
        }

        // Save the document to disk.
        doc.Save("SharedListExample.docx");
    }
}
