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

        // Add a list that uses the maximum allowed 9 levels.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        for (int level = 0; level < 9; level++)
        {
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Level {level + 1}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Validate that no list exceeds nine levels.
        foreach (List lst in doc.Lists)
        {
            int levelCount = lst.ListLevels.Count;
            if (levelCount > 9)
            {
                Console.WriteLine($"List ID {lst.ListId} exceeds the maximum of 9 levels (has {levelCount}).");
            }
            else
            {
                Console.WriteLine($"List ID {lst.ListId} is valid with {levelCount} level(s).");
            }
        }

        // Save the document (optional, just to demonstrate saving).
        doc.Save("ValidatedLists.docx");
    }
}
