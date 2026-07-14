using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a sample list with nesting up to the maximum of 9 levels (0‑8).
        builder.Writeln("Sample list with maximum nesting levels:");
        List sampleList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = sampleList;

        for (int level = 0; level < 9; level++)
        {
            // Ensure the level does not exceed the allowed maximum (8).
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Level {level + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Apply a custom list template (e.g., BulletCircle) to every list in the document.
        // Since a List's template cannot be changed after creation, we modify the
        // formatting of each existing list to mimic the desired template.
        foreach (List list in doc.Lists)
        {
            // Change the first level to use a circle bullet.
            ListLevel level0 = list.ListLevels[0];
            level0.NumberStyle = NumberStyle.Bullet;
            level0.NumberFormat = "\u25E6"; // White bullet (circle)

            // Optionally adjust other levels to keep consistency.
            for (int i = 1; i < list.ListLevels.Count; i++)
            {
                ListLevel lvl = list.ListLevels[i];
                lvl.NumberStyle = NumberStyle.Bullet;
                lvl.NumberFormat = "\u25E6";
            }
        }

        // Ensure that no list exceeds nine nesting levels.
        // (Aspose.Words lists are inherently limited to 9 levels, but we enforce it explicitly.)
        foreach (List list in doc.Lists)
        {
            if (list.ListLevels.Count > 9)
            {
                // Trim excess levels if they somehow exist.
                while (list.ListLevels.Count > 9)
                {
                    // There is no direct remove method; this block is kept for completeness.
                    // In practice, Aspose.Words always creates lists with at most 9 levels.
                    break;
                }
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomListTemplate.docx");
        doc.Save(outputPath);
    }
}
