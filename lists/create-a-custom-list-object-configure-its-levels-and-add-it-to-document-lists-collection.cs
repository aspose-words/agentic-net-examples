using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom list based on a numbered template.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure first level (1., 2., 3.).
        ListLevel level1 = customList.ListLevels[0];
        level1.NumberStyle = NumberStyle.Arabic;
        level1.NumberFormat = "%1.";
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = 0;
        level1.TextPosition = 14;
        level1.Font.Size = 12;

        // Configure second level (a), b), c)).
        ListLevel level2 = customList.ListLevels[1];
        level2.NumberStyle = NumberStyle.LowercaseLetter; // Fixed enum name
        level2.NumberFormat = "%2)";
        level2.Alignment = ListLevelAlignment.Left;
        level2.NumberPosition = 14;
        level2.TextPosition = 28;
        level2.Font.Size = 12;

        // Apply the custom list to the builder.
        builder.ListFormat.List = customList;

        // Write first‑level items.
        builder.ListFormat.ListLevelNumber = 0; // Level 1
        builder.Writeln("First item");
        builder.Writeln("Second item");

        // Write nested second‑level items.
        builder.ListFormat.ListLevelNumber = 1; // Level 2
        builder.Writeln("Nested item A");
        builder.Writeln("Nested item B");

        // Return to first level and add another item.
        builder.ListFormat.ListLevelNumber = 0; // Level 1 again
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("CustomList.docx");
    }
}
