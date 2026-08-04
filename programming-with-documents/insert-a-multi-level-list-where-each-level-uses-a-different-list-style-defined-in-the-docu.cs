using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -----------------------------------------------------------------
        // Define a multi‑level list style where each level has its own format.
        // -----------------------------------------------------------------
        Style multiLevelStyle = doc.Styles.Add(StyleType.List, "MyMultiLevelStyle");
        List listDefinition = multiLevelStyle.List; // This is the list definition belonging to the style.

        // Level 0 – bullet list.
        ListLevel level0 = listDefinition.ListLevels[0];
        level0.NumberStyle = NumberStyle.Bullet;
        level0.NumberFormat = "\u2022"; // •
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;
        level0.Font.Color = Color.DarkBlue;
        level0.NumberPosition = -18;
        level0.TextPosition = 18;
        level0.TabPosition = 36;

        // Level 1 – decimal numbers (1., 2., …).
        ListLevel level1 = listDefinition.ListLevels[1];
        level1.NumberStyle = NumberStyle.Arabic;
        level1.NumberFormat = "%1.";
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        level1.Font.Color = Color.DarkGreen;
        level1.NumberPosition = -36;
        level1.TextPosition = 36;
        level1.TabPosition = 72;

        // Level 2 – lower‑case letters (a., b., …).
        ListLevel level2 = listDefinition.ListLevels[2];
        level2.NumberStyle = NumberStyle.LowercaseLetter;
        level2.NumberFormat = "%2.";
        level2.Font.Name = "Arial";
        level2.Font.Size = 12;
        level2.Font.Color = Color.Maroon;
        level2.NumberPosition = -54;
        level2.TextPosition = 54;
        level2.TabPosition = 108;

        // -----------------------------------------------------------------
        // Use the defined list style in the document.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list that references the style we just defined.
        List multiLevelList = doc.Lists.Add(multiLevelStyle);
        builder.ListFormat.List = multiLevelList;

        // First level item.
        builder.Writeln("Level 1 – Item 1");

        // Second level item.
        builder.ListFormat.ListIndent();
        builder.Writeln("Level 2 – Item 1");

        // Third level item.
        builder.ListFormat.ListIndent();
        builder.Writeln("Level 3 – Item 1");

        // Back to second level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Level 2 – Item 2");

        // Back to first level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Level 1 – Item 2");

        // Finish the list.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "MultiLevelList.docx");
        doc.Save(outputPath);
    }
}
