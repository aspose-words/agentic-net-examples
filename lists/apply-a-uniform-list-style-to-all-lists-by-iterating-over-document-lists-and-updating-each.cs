using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class ApplyUniformListStyle
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Get the document builder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a bulleted list.
        builder.ListFormat.ApplyBulletDefault(); // use bullet list format
        builder.Writeln("Bullet item 1");
        builder.Writeln("Bullet item 2");
        builder.ListFormat.RemoveNumbers();

        // Add a numbered list.
        builder.ListFormat.ApplyNumberDefault(); // use numbered list format
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");
        builder.ListFormat.RemoveNumbers();

        // Iterate over all lists in the document.
        foreach (List list in doc.Lists)
        {
            // Iterate over each level of the list (typically 0-8).
            for (int i = 0; i < list.ListLevels.Count; i++)
            {
                ListLevel level = list.ListLevels[i];

                // Apply a uniform font to the list level.
                level.Font.Name = "Arial";
                level.Font.Size = 12;

                // Ensure the list level aligns left.
                level.Alignment = ListLevelAlignment.Left;

                // Optional: set a uniform tab position.
                level.TabPosition = 36;
            }
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
