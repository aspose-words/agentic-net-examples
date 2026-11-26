using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Create a numbered list ----------
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First numbered item");
        builder.Writeln("Second numbered item");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // ---------- Create a bullet list ----------
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // ---------- Iterate through all lists and modify them uniformly ----------
        foreach (List list in doc.Lists)
        {
            foreach (ListLevel level in list.ListLevels)
            {
                // Set a common font for list numbers/bullets.
                level.Font.Name = "Arial";
                level.Font.Size = 12;

                // Apply a uniform numbering style.
                if (level.NumberStyle == NumberStyle.Bullet)
                {
                    // Ensure bullet lists use a standard bullet character.
                    level.NumberStyle = NumberStyle.Bullet;
                    level.NumberFormat = "\u2022"; // •
                }
                else
                {
                    // Use Uppercase Roman numerals for numbered lists.
                    level.NumberStyle = NumberStyle.UppercaseRoman;
                }
            }
        }

        // Save the modified document.
        doc.Save("ModifiedLists.docx");
    }
}
