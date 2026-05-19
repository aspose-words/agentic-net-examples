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

        // Start a default numbered list.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add a first‑level list item (level 0).
        builder.Writeln("Level 0 item");

        // Set the list level to 3 (third‑level item, zero‑based index).
        builder.ListFormat.ListLevelNumber = 3;
        builder.Writeln("Third‑level list item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the file system.
        doc.Save("ThirdLevelList.docx");
    }
}
