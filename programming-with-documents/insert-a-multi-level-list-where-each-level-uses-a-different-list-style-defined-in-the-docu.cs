using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -----------------------------------------------------------------
        // Define a custom list style that will be used for all three levels.
        // Each level gets its own formatting (bullet, lower‑case letter,
        // lower‑case Roman numeral).
        // -----------------------------------------------------------------
        Style multiLevelStyle = doc.Styles.Add(StyleType.List, "MyMultiLevelList");

        // The List object that belongs to the style – this is the definition.
        List listDefinition = multiLevelStyle.List;

        // Level 0 – bullet (disc).
        ListLevel level0 = listDefinition.ListLevels[0];
        level0.NumberStyle = NumberStyle.Bullet;
        level0.NumberFormat = "\u2022";               // •
        level0.Font.Name = "Symbol";
        level0.Font.Size = 12;
        level0.Font.Color = Color.Black;
        level0.NumberPosition = -18;                  // bullet left of text
        level0.TextPosition = 18;                     // text after bullet
        level0.TabPosition = 36;

        // Level 1 – lower‑case letters (a., b., …).
        ListLevel level1 = listDefinition.ListLevels[1];
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.NumberFormat = "%1.";                  // a., b., …
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        level1.Font.Color = Color.DarkBlue;
        level1.NumberPosition = -36;
        level1.TextPosition = 36;
        level1.TabPosition = 72;

        // Level 2 – lower‑case Roman numerals (i., ii., …).
        ListLevel level2 = listDefinition.ListLevels[2];
        level2.NumberStyle = NumberStyle.LowercaseRoman;
        level2.NumberFormat = "%1.";                  // i., ii., …
        level2.Font.Name = "Times New Roman";
        level2.Font.Size = 12;
        level2.Font.Color = Color.DarkGreen;
        level2.NumberPosition = -54;
        level2.TextPosition = 54;
        level2.TabPosition = 108;

        // -----------------------------------------------------------------
        // Create a list that references the style defined above.
        // -----------------------------------------------------------------
        List multiLevelList = doc.Lists.Add(multiLevelStyle);

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the list to the builder.
        builder.ListFormat.List = multiLevelList;

        // Level 0 items.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Level 1 items.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");

        // Level 2 items.
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Third level item 1");
        builder.Writeln("Third level item 2");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local folder.
        string outputPath = "MultiLevelList.docx";
        doc.Save(outputPath);
    }
}
