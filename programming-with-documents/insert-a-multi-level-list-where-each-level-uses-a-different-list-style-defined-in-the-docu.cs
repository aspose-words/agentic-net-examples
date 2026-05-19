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

        // -----------------------------------------------------------------
        // Define three separate list styles – one for each level.
        // -----------------------------------------------------------------
        // Level 1 style – decimal numbers (1., 2., 3.)
        Style level1Style = doc.Styles.Add(StyleType.List, "Level1Style");
        List level1Def = level1Style.List; // Definition of the list style.
        level1Def.ListLevels[0].NumberStyle = NumberStyle.Arabic;
        level1Def.ListLevels[0].NumberFormat = "%1.";
        level1Def.ListLevels[0].Font.Name = "Calibri";
        level1Def.ListLevels[0].Font.Size = 12;

        // Level 2 style – lower‑case letters (a., b., c.)
        Style level2Style = doc.Styles.Add(StyleType.List, "Level2Style");
        List level2Def = level2Style.List;
        level2Def.ListLevels[0].NumberStyle = NumberStyle.LowercaseLetter;
        level2Def.ListLevels[0].NumberFormat = "%1.";
        level2Def.ListLevels[0].Font.Name = "Calibri";
        level2Def.ListLevels[0].Font.Size = 12;

        // Level 3 style – bullet (Wingdings star)
        Style level3Style = doc.Styles.Add(StyleType.List, "Level3Style");
        List level3Def = level3Style.List;
        level3Def.ListLevels[0].NumberStyle = NumberStyle.Bullet;
        level3Def.ListLevels[0].NumberFormat = "\uF0AF"; // Star symbol.
        level3Def.ListLevels[0].Font.Name = "Wingdings";
        level3Def.ListLevels[0].Font.Size = 12;

        // -----------------------------------------------------------------
        // Create list instances that reference the styles defined above.
        // -----------------------------------------------------------------
        List listLevel1 = doc.Lists.Add(level1Style);
        List listLevel2 = doc.Lists.Add(level2Style);
        List listLevel3 = doc.Lists.Add(level3Style);

        // -----------------------------------------------------------------
        // Build a multi‑level list where each level uses its own style.
        // -----------------------------------------------------------------
        // First level item.
        builder.ListFormat.List = listLevel1;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item");

        // Second level item.
        builder.ListFormat.List = listLevel2;
        builder.ListFormat.ListLevelNumber = 1; // Indent to level 2.
        builder.Writeln("Second level item");

        // Third level item.
        builder.ListFormat.List = listLevel3;
        builder.ListFormat.ListLevelNumber = 2; // Indent to level 3.
        builder.Writeln("Third level item");

        // Add another first‑level item to demonstrate continuation.
        builder.ListFormat.List = listLevel1;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item (second)");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // Save the document to the current directory.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MultiLevelList.docx");
        doc.Save(outputPath);
    }
}
