using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListIndentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.ApplyNumberDefault();

            // Write the first list item (level 0).
            builder.Writeln("Item at level 0");

            // Increase the list level inside a loop and add items at each deeper level.
            for (int i = 1; i <= 5; i++)
            {
                // Increase the current list level by one.
                builder.ListFormat.ListIndent();

                // Write a paragraph that will appear at the new list level.
                builder.Writeln($"Item at level {i}");
            }

            // Optional: return to the original level.
            for (int i = 0; i < 5; i++)
                builder.ListFormat.ListOutdent();

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Prepare an output folder.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "IncreaseIndent.docx");
            doc.Save(outputPath);
        }
    }
}
