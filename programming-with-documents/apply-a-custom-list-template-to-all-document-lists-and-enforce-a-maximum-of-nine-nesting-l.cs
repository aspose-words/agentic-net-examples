using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a list with more than nine nesting levels.
        builder.Writeln("Start of document");
        builder.ListFormat.ApplyNumberDefault();
        for (int i = 1; i <= 12; i++)
        {
            builder.Writeln($"Level {i}");
            if (i < 12)
                builder.ListFormat.ListIndent();
        }
        // Close the list.
        builder.ListFormat.RemoveNumbers();

        // Apply a custom list template to all lists in the document.
        foreach (List list in doc.Lists)
        {
            // Define a simple custom bullet for the first level.
            ListLevel level = list.ListLevels[0];
            // Set bullet style via NumberFormat; omit NumberStyle to avoid missing enum.
            level.Font.Name = "Arial";
            level.Font.Size = 10;
            level.Font.Color = Color.Blue;
            level.NumberFormat = "•";

            // Ensure remaining levels inherit the same style.
            for (int lvl = 1; lvl < list.ListLevels.Count; lvl++)
            {
                ListLevel src = list.ListLevels[0];
                ListLevel dst = list.ListLevels[lvl];
                dst.Font.Name = src.Font.Name;
                dst.Font.Size = src.Font.Size;
                dst.Font.Color = src.Font.Color;
                dst.NumberFormat = src.NumberFormat;
            }
        }

        // Enforce a maximum of nine nesting levels.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.IsListItem && para.ListFormat.ListLevelNumber > 9)
            {
                para.ListFormat.ListLevelNumber = 9;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Verify the file was saved.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
