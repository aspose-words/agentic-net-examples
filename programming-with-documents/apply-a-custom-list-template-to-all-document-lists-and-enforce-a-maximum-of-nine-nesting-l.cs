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

        // Insert sample list items with varying nesting levels (including deeper than 9).
        for (int i = 0; i < 12; i++)
        {
            // Start a numbered list using the default list template.
            builder.ListFormat.ApplyNumberDefault();

            // Set the desired list level (0‑based). Levels 0‑10 correspond to nesting levels 1‑11.
            builder.ListFormat.ListLevelNumber = i; // Intentionally exceed the max of 8 (level 9).

            // Write the list item.
            builder.Writeln($"List item at intended level {i + 1}");
        }

        // End the list formatting for subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Create a custom list template (Numbered list with custom formatting).
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Example: customize the first three levels.
        for (int level = 0; level < 3; level++)
        {
            ListLevel listLevel = customList.ListLevels[level];
            listLevel.NumberFormat = $"{level + 1}.";
            listLevel.NumberStyle = NumberStyle.Arabic;
            listLevel.Alignment = ListLevelAlignment.Left;
            listLevel.Font.Size = 12;
            listLevel.Font.Name = "Arial";
        }

        // Apply the custom list to all existing list items and enforce a maximum nesting level of 9.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.IsListItem)
            {
                // Assign the custom list.
                para.ListFormat.List = customList;

                // Enforce maximum nesting level of 9 (zero‑based index 8).
                if (para.ListFormat.ListLevelNumber > 8)
                {
                    para.ListFormat.ListLevelNumber = 8;
                }
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomListTemplate.docx");
        doc.Save(outputPath);

        // Verify the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
