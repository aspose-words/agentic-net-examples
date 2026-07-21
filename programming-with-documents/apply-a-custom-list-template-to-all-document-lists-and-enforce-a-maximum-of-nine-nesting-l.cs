using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListTemplateExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build sample lists with varying nesting levels (up to 9 levels).
            // First list – numbered.
            builder.Writeln("Sample numbered list:");
            builder.ListFormat.ApplyNumberDefault(); // start default numbered list

            for (int level = 0; level < 9; level++)
            {
                // Ensure the level does not exceed the maximum (0‑8).
                builder.ListFormat.ListLevelNumber = level;
                builder.Writeln($"Level {level + 1}");
            }

            builder.ListFormat.RemoveNumbers(); // end the first list.

            // Second list – bulleted.
            builder.Writeln();
            builder.Writeln("Sample bulleted list:");
            builder.ListFormat.ApplyBulletDefault(); // start default bulleted list

            for (int level = 0; level < 9; level++)
            {
                builder.ListFormat.ListLevelNumber = level;
                builder.Writeln($"Bullet level {level + 1}");
            }

            builder.ListFormat.RemoveNumbers(); // end the second list.

            // -----------------------------------------------------------------
            // Apply a custom list template to all existing list items.
            // The chosen template is NumberUppercaseLetterDot (A., B., C., …).
            // -----------------------------------------------------------------
            List customList = doc.Lists.Add(ListTemplate.NumberUppercaseLetterDot);

            // Iterate over every paragraph in the document.
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
            {
                // If the paragraph is part of any list, reassign it to the custom list.
                if (para.ListFormat.IsListItem)
                {
                    // Clamp the level to the allowed range (0‑8) to enforce a maximum of nine nesting levels.
                    int level = para.ListFormat.ListLevelNumber;
                    if (level > 8)
                        level = 8;

                    para.ListFormat.List = customList;
                    para.ListFormat.ListLevelNumber = level;
                }
            }

            // Save the resulting document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomListTemplate.docx");
            doc.Save(outputPath);
        }
    }
}
