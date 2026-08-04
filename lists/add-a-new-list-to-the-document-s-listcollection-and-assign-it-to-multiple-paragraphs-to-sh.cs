using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list to the document's ListCollection using a predefined template.
        List sharedList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Apply the same list to multiple paragraphs.
        builder.ListFormat.List = sharedList;
        builder.ListFormat.ListLevelNumber = 0; // Use the first level of the list.

        builder.Writeln("First shared list item");
        builder.Writeln("Second shared list item");
        builder.Writeln("Third shared list item");

        // Stop list formatting for any following paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Add a regular paragraph without list formatting.
        builder.Writeln("A normal paragraph without list formatting.");

        // Save the document to disk.
        doc.Save("SharedListExample.docx");
    }
}
