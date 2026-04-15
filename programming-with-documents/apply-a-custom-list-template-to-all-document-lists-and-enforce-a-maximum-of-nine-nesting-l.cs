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

        // Create a custom list using the BulletCircle template.
        List customList = doc.Lists.Add(ListTemplate.BulletCircle);

        // Apply the custom list to paragraphs, limiting nesting to nine levels.
        for (int level = 0; level < 9; level++)
        {
            builder.ListFormat.List = customList;
            builder.ListFormat.ListLevelNumber = level; // 0‑based level index.
            builder.Writeln($"Level {level + 1}");
        }

        // End the first list.
        builder.ListFormat.RemoveNumbers();

        // Add another list to show that all document lists use the same template.
        List secondList = doc.Lists.Add(ListTemplate.BulletCircle);
        for (int level = 0; level < 9; level++)
        {
            builder.ListFormat.List = secondList;
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Second list level {level + 1}");
        }
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        const string outputFile = "CustomListTemplate.docx";
        doc.Save(outputFile);
    }
}
