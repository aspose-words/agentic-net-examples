using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ListExample
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Load the document using the Document(string) constructor (lifecycle rule).
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Add a new list based on a predefined template.
        // ------------------------------------------------------------
        // Use ListCollection.Add(ListTemplate) (feature rule) to create a numbered list.
        List newList = doc.Lists.Add(ListTemplate.NumberDefault);

        // ------------------------------------------------------------
        // 2. Apply the list to newly inserted paragraphs.
        // ------------------------------------------------------------
        builder.Writeln("=== New List Begins ===");

        // Set the list for the builder – this makes subsequent paragraphs list items.
        builder.ListFormat.List = newList;

        // First level item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 1");

        // Second level (sub‑item) – increase the level number.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Sub‑item 1.1");

        // Back to first level.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 2");

        // ------------------------------------------------------------
        // 3. End list formatting.
        // ------------------------------------------------------------
        builder.ListFormat.RemoveNumbers();

        // ------------------------------------------------------------
        // 4. (Optional) Enumerate all lists in the document and output their IDs.
        // ------------------------------------------------------------
        Console.WriteLine("Document contains {0} list(s):", doc.Lists.Count);
        foreach (List list in doc.Lists)
        {
            Console.WriteLine(" - ListId: {0}", list.ListId);
        }

        // ------------------------------------------------------------
        // 5. Save the modified document using the Document.Save(string) method (lifecycle rule).
        // ------------------------------------------------------------
        string outputPath = @"C:\Docs\OutputDocument.docx";
        doc.Save(outputPath);
    }
}
