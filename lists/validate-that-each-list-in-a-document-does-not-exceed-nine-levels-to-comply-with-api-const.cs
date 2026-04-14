using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a list with the default numbered template (contains 9 levels).
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add some paragraphs using the list to demonstrate the list exists.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        for (int i = 0; i < 3; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // use different levels (0‑2)
            builder.Writeln($"Item at level {i}");
        }
        builder.ListFormat.RemoveNumbers();

        // Validate that each list in the document does not exceed nine levels.
        bool allValid = true;
        foreach (List lst in doc.Lists)
        {
            int levelCount = lst.ListLevels.Count; // ListLevelCollection.Count
            if (levelCount > 9)
            {
                allValid = false;
                Console.WriteLine($"List with ID {lst.ListId} has {levelCount} levels, which exceeds the allowed maximum of 9.");
            }
            else
            {
                Console.WriteLine($"List with ID {lst.ListId} has {levelCount} levels – within the allowed range.");
            }
        }

        // Output overall validation result.
        Console.WriteLine(allValid
            ? "All lists are within the allowed level limit."
            : "One or more lists exceed the allowed level limit.");

        // Save the document (output file path can be adjusted as needed).
        doc.Save("ValidatedLists.docx");
    }
}
