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
            builder.Writeln("Root level item");

            // Increase the list level inside a loop to create nested items.
            for (int i = 1; i <= 5; i++)
            {
                // Increase the indent (list level) by one.
                builder.ListFormat.ListIndent();

                // Write a paragraph at the current list level.
                builder.Writeln($"Nested level {i}");
            }

            // Optional: return to the original level.
            for (int i = 0; i < 5; i++)
            {
                builder.ListFormat.ListOutdent();
            }

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ListIndentExample.docx");
            doc.Save(outputPath);
        }
    }
}
