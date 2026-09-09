using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a sample list (all Aspose.Words lists have up to 9 levels).
        List sampleList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Populate the list with items on each level to illustrate the structure.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = sampleList;
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // Levels are 0‑8.
            builder.Writeln($"Level {i}");
        }
        builder.ListFormat.RemoveNumbers();

        // Validate that every list in the document contains no more than nine levels.
        foreach (List list in doc.Lists)
        {
            int levelCount = list.ListLevels.Count; // Gets the number of levels in this list.
            if (levelCount > 9)
            {
                Console.WriteLine($"List ID {list.ListId} exceeds nine levels: {levelCount}");
            }
            else
            {
                Console.WriteLine($"List ID {list.ListId} is valid with {levelCount} levels.");
            }
        }

        // Save the document to disk.
        doc.Save("ValidatedLists.docx");
    }
}
