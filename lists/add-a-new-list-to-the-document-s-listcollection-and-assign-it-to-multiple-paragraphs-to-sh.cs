using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list to the document's ListCollection using a predefined template.
        // This list will be shared by multiple paragraphs.
        List sharedList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Write a normal paragraph (no list formatting).
        builder.Writeln("Paragraph without list 1");

        // Apply the shared list to subsequent paragraphs.
        builder.ListFormat.List = sharedList;
        builder.ListFormat.ListLevelNumber = 0; // Use the first level of the list.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Stop list formatting for following paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Write another normal paragraph.
        builder.Writeln("Paragraph without list 2");

        // Write paragraphs that will later be assigned the shared list manually.
        builder.Writeln("Manual paragraph 1");
        builder.Writeln("Manual paragraph 2");

        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Assign the shared list to the manually created paragraphs.
        foreach (Paragraph para in paragraphs.OfType<Paragraph>()
                                            .Where(p => p.GetText().Trim().StartsWith("Manual")))
        {
            para.ListFormat.List = sharedList;
            para.ListFormat.ListLevelNumber = 0;
        }

        // Save the document to a file in the current directory.
        doc.Save("SharedListExample.docx");
    }
}
