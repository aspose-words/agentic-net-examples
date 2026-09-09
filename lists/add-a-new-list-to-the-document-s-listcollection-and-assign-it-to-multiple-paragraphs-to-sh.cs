using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list to the document's ListCollection using a predefined template.
        List sharedList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Write a normal paragraph (no list formatting).
        builder.Writeln("Paragraph without list.");

        // Add several paragraphs that share the same list formatting.
        for (int i = 1; i <= 3; i++)
        {
            // Assign the shared list to the current paragraph.
            builder.ListFormat.List = sharedList;
            // Use the first level of the list (level 0).
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln($"List item {i}");
        }

        // Stop applying list formatting for any following paragraphs.
        builder.ListFormat.List = null;
        builder.Writeln("Paragraph after list.");

        // Save the document to a file in the current directory.
        doc.Save("SharedListExample.docx");
    }
}
