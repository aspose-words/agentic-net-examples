using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -------------------------
        // Define the first list style (bulleted).
        // -------------------------
        Style bulletListStyle = doc.Styles.Add(StyleType.List, "BulletListStyle");
        List bulletList = bulletListStyle.List; // This is the definition of the list style.

        // Configure level 0 of the bullet list.
        ListLevel bulletLevel = bulletList.ListLevels[0];
        bulletLevel.NumberStyle = NumberStyle.Bullet;
        bulletLevel.Font.Name = "Wingdings";
        bulletLevel.Font.Color = Color.DarkGreen;
        bulletLevel.NumberFormat = "\u2022"; // Standard bullet.
        bulletLevel.TrailingCharacter = ListTrailingCharacter.Tab;
        bulletLevel.NumberPosition = -18;
        bulletLevel.TextPosition = 18;
        bulletLevel.TabPosition = 36;

        // -------------------------
        // Define the second list style (numbered).
        // -------------------------
        Style numberListStyle = doc.Styles.Add(StyleType.List, "NumberListStyle");
        List numberList = numberListStyle.List;

        // Configure level 0 of the numbered list.
        ListLevel numberLevel = numberList.ListLevels[0];
        numberLevel.NumberStyle = NumberStyle.Arabic; // Use Arabic numbering instead of non‑existent Decimal.
        numberLevel.Font.Name = "Arial";
        numberLevel.Font.Color = Color.DarkBlue;
        numberLevel.NumberFormat = "%1.";
        numberLevel.TrailingCharacter = ListTrailingCharacter.Tab;
        numberLevel.NumberPosition = -18;
        numberLevel.TextPosition = 18;
        numberLevel.TabPosition = 36;

        // -------------------------
        // Create paragraph styles that reference the list styles.
        // -------------------------
        Style level1ParagraphStyle = doc.Styles.Add(StyleType.Paragraph, "Level1ParagraphStyle");
        // Attach the bullet list style to this paragraph style.
        level1ParagraphStyle.ListFormat.List = doc.Lists.Add(bulletListStyle);
        level1ParagraphStyle.ListFormat.ListLevelNumber = 0;

        Style level2ParagraphStyle = doc.Styles.Add(StyleType.Paragraph, "Level2ParagraphStyle");
        // Attach the numbered list style to this paragraph style.
        level2ParagraphStyle.ListFormat.List = doc.Lists.Add(numberListStyle);
        level2ParagraphStyle.ListFormat.ListLevelNumber = 0;

        // -------------------------
        // Build the multi‑level list using the paragraph styles.
        // -------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First top‑level item (uses bullet style).
        builder.ParagraphFormat.Style = level1ParagraphStyle;
        builder.Writeln("Top level item 1");

        // Sub‑item (uses numbered style).
        builder.ParagraphFormat.Style = level2ParagraphStyle;
        builder.Writeln("Sub‑item 1.1");

        // Another sub‑item.
        builder.Writeln("Sub‑item 1.2");

        // Back to top‑level item.
        builder.ParagraphFormat.Style = level1ParagraphStyle;
        builder.Writeln("Top level item 2");

        // Sub‑item under the second top‑level item.
        builder.ParagraphFormat.Style = level2ParagraphStyle;
        builder.Writeln("Sub‑item 2.1");

        // -------------------------
        // Save the document.
        // -------------------------
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "MultiLevelListWithDifferentStyles.docx");
        doc.Save(outputPath);
    }
}
