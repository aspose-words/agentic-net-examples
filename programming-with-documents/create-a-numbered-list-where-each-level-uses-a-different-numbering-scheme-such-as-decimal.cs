using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure level 0 to use decimal numbering.
        ListLevel level0 = list.ListLevels[0];
        level0.NumberFormat = "%1.";
        level0.NumberStyle = NumberStyle.Arabic;
        level0.NumberPosition = 0;
        level0.Alignment = ListLevelAlignment.Left;
        level0.TextPosition = 72; // 0.5 inch

        // Configure level 1 to use lower‑roman numbering.
        ListLevel level1 = list.ListLevels[1];
        level1.NumberFormat = "%2.";
        level1.NumberStyle = NumberStyle.LowercaseRoman;
        level1.NumberPosition = 0;
        level1.Alignment = ListLevelAlignment.Left;
        level1.TextPosition = 144; // 1 inch

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Level 0 item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First item (decimal)");

        // Level 1 sub‑item.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("First sub‑item (lower‑roman)");

        // Back to level 0.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Second item (decimal)");

        // End the list.
        builder.ListFormat.List = null;

        // Save the document.
        string outputPath = "NumberedList.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
