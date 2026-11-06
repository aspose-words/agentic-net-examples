using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

namespace ListLevelValidationExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list using a built‑in template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Build a list with varying nesting levels (including an invalid level > 9).
            for (int level = 1; level <= 10; level++)
            {
                // Assign the list to the builder.
                builder.ListFormat.List = list;

                // Set the desired nesting level.
                builder.ListFormat.ListLevelNumber = level;

                // Write the list item.
                builder.Writeln($"Item at level {level}");
            }

            // End list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document (optional, just to demonstrate full workflow).
            doc.Save("ListLevelValidationOutput.docx");

            // Validate that no list exceeds nine levels.
            List<Paragraph> invalidParagraphs = new List<Paragraph>();

            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.IsListItem)
                {
                    int currentLevel = para.ListFormat.ListLevelNumber;
                    if (currentLevel > 9)
                    {
                        invalidParagraphs.Add(para);
                    }
                }
            }

            // Output validation results.
            if (invalidParagraphs.Count == 0)
            {
                Console.WriteLine("All lists are within the allowed nine levels.");
            }
            else
            {
                Console.WriteLine($"Found {invalidParagraphs.Count} paragraph(s) exceeding nine list levels:");
                foreach (Paragraph p in invalidParagraphs)
                {
                    Console.WriteLine($"- Text: \"{p.GetText().Trim()}\" (Level {p.ListFormat.ListLevelNumber})");
                }
            }
        }
    }
}
