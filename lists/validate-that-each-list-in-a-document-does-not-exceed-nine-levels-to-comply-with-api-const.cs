using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a sample list that has the maximum allowed 9 levels.
        // All lists created via Add(ListTemplate) contain 9 levels by definition.
        List sampleList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Optionally add some paragraphs to demonstrate the list (not required for validation).
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = sampleList;
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // Set level 0‑8.
            builder.Writeln($"Level {i + 1}");
        }
        builder.ListFormat.RemoveNumbers();

        // Validate that no list exceeds nine levels.
        foreach (List list in doc.Lists)
        {
            int levelCount = list.ListLevels.Count; // Gets the number of levels in this list.
            if (levelCount > 9)
            {
                // This should never happen with Aspose.Words, but we report it just in case.
                Console.WriteLine($"List with ID {list.ListId} has {levelCount} levels, which exceeds the allowed maximum of 9.");
            }
            else
            {
                Console.WriteLine($"List with ID {list.ListId} is valid with {levelCount} levels.");
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ValidatedLists.docx");
        doc.Save(outputPath);
    }
}
