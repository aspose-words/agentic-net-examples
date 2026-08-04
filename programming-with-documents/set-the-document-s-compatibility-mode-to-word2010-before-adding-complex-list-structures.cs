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

        // Set the document's compatibility mode to Word 2010.
        // This prevents Word from showing the "Compatibility mode" ribbon.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Create a multilevel list based on a predefined template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Customize the first three list levels (optional, demonstrates complex structure).
        ListLevel level0 = list.ListLevels[0];
        level0.NumberFormat = "%1.";
        level0.NumberStyle = NumberStyle.Arabic;
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;

        ListLevel level1 = list.ListLevels[1];
        level1.NumberFormat = "%2.";
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;

        ListLevel level2 = list.ListLevels[2];
        level2.NumberFormat = "%3.";
        level2.NumberStyle = NumberStyle.LowercaseRoman;
        level2.Font.Name = "Arial";
        level2.Font.Size = 12;

        // Use DocumentBuilder to add list items with varying levels.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // First level item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 1");

        // Second level item.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Second level item 1");

        // Third level item.
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Third level item 1");

        // Back to first level.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First level item 2");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        doc.Save("ComplexList.docx");
    }
}
