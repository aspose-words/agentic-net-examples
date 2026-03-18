using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ListLevelExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a default numbered list to the builder.
        // This creates a list object that the builder can work with.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the list level to the third level.
        // ListLevelNumber is zero‑based (0 = first level), so 2 corresponds to the third level.
        builder.ListFormat.ListLevelNumber = 2;

        // Add a paragraph that will appear as a third‑level list item.
        builder.Writeln("Third‑level list item");

        // Remove list formatting if further paragraphs are added (optional).
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ThirdLevelListItem.docx");
    }
}
