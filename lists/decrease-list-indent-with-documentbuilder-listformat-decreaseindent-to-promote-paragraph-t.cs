using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListOutdentExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1 (level 0)");
            builder.Writeln("Item 2 (level 0)");

            // Increase the list level – this creates a sub‑list.
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 1 (level 1)");
            builder.Writeln("Item 2 (level 1)");

            // Decrease the list level – the next paragraphs return to the previous level.
            builder.ListFormat.ListOutdent();
            builder.Writeln("Item 3 (back to level 0)");
            builder.Writeln("Item 4 (level 0)");

            // End the list.
            builder.ListFormat.RemoveNumbers();

            // Prepare an output folder.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "ListOutdentDemo.docx");
            doc.Save(outputPath);
        }
    }
}
