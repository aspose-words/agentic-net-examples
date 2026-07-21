using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Optimize the document for Word 2010 to avoid the Compatibility mode ribbon.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Create a multi‑level numbered list based on a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Customize the first level (1., 2., 3., ...).
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;
        level0.NumberStyle = NumberStyle.Arabic;          // Fixed: use Arabic instead of Decimal
        level0.NumberFormat = "%1.";
        level0.Alignment = ListLevelAlignment.Left;
        level0.NumberPosition = -18;   // Indent for the number.
        level0.TextPosition = 0;       // Position of the text.
        level0.TabPosition = 36;       // Tab stop after the number.

        // Customize the second level (a), b), c), ...).
        ListLevel level1 = list.ListLevels[1];
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        level1.NumberStyle = NumberStyle.LowercaseLetter; // Fixed: use LowercaseLetter instead of LowerLetter
        level1.NumberFormat = "%2)";
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = 18;
        level1.TextPosition = 36;
        level1.TabPosition = 72;

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // First‑level items.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Second‑level items.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");

        // Back to first level.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 3");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComplexList.docx");
        doc.Save(outputPath);
    }
}
