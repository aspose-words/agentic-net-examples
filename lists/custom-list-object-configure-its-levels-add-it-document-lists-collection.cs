using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class CustomListExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new list based on a predefined template.
        // The template provides a multi‑level list (9 levels) that we can customize.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // ----- Configure first level (level 0) -----
        ListLevel level0 = customList.ListLevels[0];
        level0.NumberStyle = NumberStyle.UppercaseLetter;   // A., B., C., …
        level0.Font.Color = Color.Blue;                     // Blue letters for the label.
        level0.StartAt = 5;                                 // Start numbering at E.
        level0.NumberFormat = "%1.";                        // Use the default format for letters.
        level0.Alignment = ListLevelAlignment.Right;        // Align the label to the right.

        // ----- Configure second level (level 1) -----
        ListLevel level1 = customList.ListLevels[1];
        level1.NumberStyle = NumberStyle.Bullet;            // Use a bullet.
        level1.NumberFormat = "\u2022";                     // Unicode bullet character.
        level1.Font.Name = "Wingdings";                     // Font that contains the bullet glyph.
        level1.Font.Color = Color.Red;                      // Red bullet.
        level1.Alignment = ListLevelAlignment.Left;         // Align the bullet to the left.

        // Use DocumentBuilder to add paragraphs that use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First level item.
        builder.ListFormat.List = customList;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item");

        // Second level (nested) item.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Second level sub‑item");

        // Return to first level for another item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Another first level item");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("CustomList.docx");
    }
}
