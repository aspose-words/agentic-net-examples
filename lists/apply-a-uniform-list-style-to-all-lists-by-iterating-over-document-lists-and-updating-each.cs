using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Numbered list ----------
        // Create a numbered list instance.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        // Assign the list to the builder.
        builder.ListFormat.List = numberedList;
        // Write list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // ---------- Bullet list with a nested level ----------
        // Create a bullet list instance.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        // First level bullet.
        builder.Writeln("Bullet item 1");
        // Switch to second level (nested).
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Nested bullet item");
        // Return to first level.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Bullet item 2");
        // End the bullet list.
        builder.ListFormat.RemoveNumbers();

        // ---------- Apply uniform formatting to all list levels ----------
        foreach (List list in doc.Lists)
        {
            foreach (ListLevel level in list.ListLevels)
            {
                // Consistent font for the list level.
                level.Font.Name = "Arial";
                level.Font.Size = 12;

                // Uniform left alignment and indentation.
                level.Alignment = ListLevelAlignment.Left;
                level.NumberPosition = 0;
                level.TextPosition = 18;
            }
        }

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
