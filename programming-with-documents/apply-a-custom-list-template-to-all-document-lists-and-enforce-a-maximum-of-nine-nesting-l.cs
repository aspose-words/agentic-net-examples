using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Sample Document with Lists");
        builder.Writeln();

        // Create a list with 10 items, but enforce a maximum of nine nesting levels.
        builder.ListFormat.ApplyNumberDefault();
        for (int level = 0; level < 10; level++)
        {
            builder.Writeln($"Level {level + 1}");
            // Indent only while we are below the ninth nesting level (index 8).
            if (level < 8)
                builder.ListFormat.ListIndent();
        }

        // Reset list formatting.
        builder.ListFormat.RemoveNumbers();

        // --------------------------------------------------------------------
        // Create a custom list template (Arabic numbers with a period) and
        // customize its first nine levels.
        // --------------------------------------------------------------------
        List customList = doc.Lists.Add(ListTemplate.NumberArabicDot);
        for (int i = 0; i < 9; i++)
        {
            ListLevel level = customList.ListLevels[i];
            level.Font.Name = "Arial";
            level.Font.Size = 10;
            level.NumberStyle = NumberStyle.Arabic;
            level.NumberFormat = $"{i + 1}.";
            level.Alignment = ListLevelAlignment.Left;
            level.NumberPosition = 0;
            level.TextPosition = 20;
            level.TabPosition = 20;
        }

        // --------------------------------------------------------------------
        // Apply the custom list template to all existing lists in the document.
        // --------------------------------------------------------------------
        foreach (List existingList in doc.Lists)
        {
            // Copy the first nine ListLevel definitions from the custom list.
            int levelsToCopy = Math.Min(9, Math.Min(customList.ListLevels.Count, existingList.ListLevels.Count));
            for (int i = 0; i < levelsToCopy; i++)
            {
                ListLevel source = customList.ListLevels[i];
                ListLevel target = existingList.ListLevels[i];

                target.Font.Name = source.Font.Name;
                target.Font.Size = source.Font.Size;
                target.NumberStyle = source.NumberStyle;
                target.NumberFormat = source.NumberFormat;
                target.Alignment = source.Alignment;
                target.NumberPosition = source.NumberPosition;
                target.TextPosition = source.TextPosition;
                target.TabPosition = source.TabPosition;
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
