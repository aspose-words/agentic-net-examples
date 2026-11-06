using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom list with two levels.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Level 1 – bullet style.
        ListLevel level1 = customList.ListLevels[0];
        // The NumberStyle enum may not be available in some library versions,
        // so we configure the level using NumberFormat and other properties.
        level1.NumberFormat = "\u2022"; // bullet character
        level1.Font.Name = "Symbol";
        level1.Font.Size = 12;
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = 0;
        level1.TextPosition = 18;
        level1.TabPosition = 36;

        // Level 2 – decimal numbering style.
        ListLevel level2 = customList.ListLevels[1];
        // Configure decimal numbering without using the NumberStyle enum.
        level2.NumberFormat = "%1.";
        level2.Alignment = ListLevelAlignment.Left;
        level2.NumberPosition = 0;
        level2.TextPosition = 36;
        level2.TabPosition = 72;

        // Insert first‑level list item.
        builder.ListFormat.List = customList;
        builder.ListFormat.ListLevelNumber = 0; // level 1
        builder.Writeln("First level item");

        // Insert second‑level list item.
        builder.ListFormat.ListLevelNumber = 1; // level 2
        builder.Writeln("Second level item");

        // Reset list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string outputPath = "MultiLevelList.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document successfully created: {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("Failed to create the document.");
        }
    }
}
