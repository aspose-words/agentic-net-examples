using System;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the document's compatibility mode to Word 2010.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Initialize a DocumentBuilder for adding content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading for the list section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Complex List Example");

        // Reset to normal style for list items.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Create a multilevel numbered list based on a predefined template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Customize first level formatting.
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;
        level0.NumberStyle = NumberStyle.Arabic;          // Decimal numbers -> Arabic
        level0.NumberFormat = "%1.";
        level0.Alignment = ListLevelAlignment.Left;
        level0.NumberPosition = -18; // Indent for the number.
        level0.TextPosition = 18;    // Indent for the text.
        level0.TabPosition = 36;

        // Customize second level formatting.
        ListLevel level1 = list.ListLevels[1];
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        level1.NumberStyle = NumberStyle.LowercaseLetter; // Lowercase letters (a, b, c...)
        level1.NumberFormat = "%2)";
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = 0;
        level1.TextPosition = 36;
        level1.TabPosition = 54;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First level items.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Second level items under "Item 2".
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
