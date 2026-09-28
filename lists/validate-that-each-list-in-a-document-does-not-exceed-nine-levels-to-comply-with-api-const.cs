using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a list with 10 levels (exceeds the nine‑level limit).
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Level 1");
        for (int i = 2; i <= 10; i++)
        {
            builder.ListFormat.ListIndent();
            builder.Writeln($"Level {i}");
        }

        // Validate that each list does not exceed nine levels.
        bool allValid = true;
        foreach (List list in doc.Lists)
        {
            // ListLevelCollection contains the defined levels for the list.
            int definedLevels = list.ListLevels.Count;
            if (definedLevels > 9)
            {
                allValid = false;
                Console.WriteLine($"List ID {list.ListId} exceeds nine levels: {definedLevels} levels.");
            }
        }

        if (allValid)
            Console.WriteLine("All lists are within nine levels.");
        else
            Console.WriteLine("One or more lists exceed nine levels.");

        // Save the document (optional, demonstrates saving).
        doc.Save("Output.docx");
    }
}
