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

        // Optimize the document for Microsoft Word 2010.
        // This sets the compatibility mode and adjusts internal flags accordingly.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Create a complex multilevel list.
        // Use a predefined template as a base.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Customize the first level (Arabic numbers).
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;
        level0.NumberStyle = NumberStyle.Arabic;
        level0.NumberFormat = "%1.";
        level0.Alignment = ListLevelAlignment.Left;
        level0.NumberPosition = -18;
        level0.TextPosition = 0;
        level0.TabPosition = 36;

        // Customize the second level (lowercase letters).
        ListLevel level1 = list.ListLevels[1];
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.NumberFormat = "%2)";
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = -36;
        level1.TextPosition = 18;
        level1.TabPosition = 54;

        // Customize the third level (lowercase Roman numerals).
        ListLevel level2 = list.ListLevels[2];
        level2.Font.Name = "Arial";
        level2.Font.Size = 12;
        level2.NumberStyle = NumberStyle.LowercaseRoman;
        level2.NumberFormat = "%3.";
        level2.Alignment = ListLevelAlignment.Left;
        level2.NumberPosition = -54;
        level2.TextPosition = 36;
        level2.TabPosition = 72;

        // Use DocumentBuilder to insert list items.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First level items.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Second level items.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");

        // Third level items.
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Third level item 1");
        builder.Writeln("Third level item 2");

        // Return to normal paragraph formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        doc.Save("ComplexList.docx");
    }
}
