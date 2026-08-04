using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list using the default list template.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the list level to 3 (zero‑based, so this is the fourth level in Word,
        // but the requirement states "three", which corresponds to level index 3).
        builder.ListFormat.ListLevelNumber = 3;

        // Add a paragraph that will appear as a third‑level list item.
        builder.Writeln("Third‑level list item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        string outputPath = "ThirdLevelList.docx";
        doc.Save(outputPath);
    }
}
