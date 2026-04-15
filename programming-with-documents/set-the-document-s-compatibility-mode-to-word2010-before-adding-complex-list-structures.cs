using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Optimize the document for Microsoft Word 2010 compatibility mode.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Create a multi‑level numbered list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level (1., 2., 3., ...).
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;
        // Use Arabic numbers for the first level.
        level0.NumberStyle = NumberStyle.Arabic;
        level0.NumberFormat = "%1.";
        level0.Alignment = ListLevelAlignment.Left;
        level0.NumberPosition = -18;   // Indent for the number.
        level0.TextPosition = 36;      // Indent for the text.
        level0.TabPosition = 36;

        // Configure the second list level (a), b), c), ...).
        ListLevel level1 = list.ListLevels[1];
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        // Use lowercase letters for the second level.
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.NumberFormat = "%2)";
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = -36;
        level1.TextPosition = 72;
        level1.TabPosition = 72;

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Complex List Example:");

        // First‑level items.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Second‑level (nested) items.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Subitem 2.1");
        builder.Writeln("Subitem 2.2");

        // Back to first level.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 3");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        doc.Save("ComplexList.docx");
    }
}
