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

        // ---------- Define three separate list styles ----------
        // Style for first level.
        Style level1Style = doc.Styles.Add(StyleType.List, "Level1Style");
        List level1ListDef = level1Style.List;
        level1ListDef.ListLevels[0].NumberStyle = NumberStyle.Arabic;          // Fixed: use Arabic instead of non‑existent Decimal
        level1ListDef.ListLevels[0].NumberFormat = "%1.";
        level1ListDef.ListLevels[0].Font.Name = "Arial";
        level1ListDef.ListLevels[0].Font.Color = System.Drawing.Color.DarkBlue;

        // Style for second level.
        Style level2Style = doc.Styles.Add(StyleType.List, "Level2Style");
        List level2ListDef = level2Style.List;
        level2ListDef.ListLevels[0].NumberStyle = NumberStyle.LowercaseLetter;
        level2ListDef.ListLevels[0].NumberFormat = "%1)";
        level2ListDef.ListLevels[0].Font.Name = "Calibri";
        level2ListDef.ListLevels[0].Font.Color = System.Drawing.Color.DarkGreen;

        // Style for third level.
        Style level3Style = doc.Styles.Add(StyleType.List, "Level3Style");
        List level3ListDef = level3Style.List;
        level3ListDef.ListLevels[0].NumberStyle = NumberStyle.Bullet;
        level3ListDef.ListLevels[0].NumberFormat = "\u2022"; // bullet character
        level3ListDef.ListLevels[0].Font.Name = "Times New Roman";
        level3ListDef.ListLevels[0].Font.Color = System.Drawing.Color.Maroon;

        // ---------- Create list instances that reference the styles ----------
        List level1List = doc.Lists.Add(level1Style);
        List level2List = doc.Lists.Add(level2Style);
        List level3List = doc.Lists.Add(level3Style);

        // Use DocumentBuilder to insert paragraphs with the different list styles.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First level items.
        builder.ListFormat.List = level1List;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Second level items.
        builder.ListFormat.List = level2List;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");

        // Third level items.
        builder.ListFormat.List = level3List;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Third level item 1");
        builder.Writeln("Third level item 2");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Ensure output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "MultiLevelListWithDifferentStyles.docx");
        doc.Save(outputPath);
    }
}
