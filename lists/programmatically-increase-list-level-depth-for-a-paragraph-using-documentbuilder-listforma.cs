using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add several paragraphs, increasing the list level each time.
        for (int i = 0; i < 5; i++)
        {
            // Write a paragraph at the current list level.
            builder.Writeln($"Item at level {i}");

            // Increase the list level for the next paragraph.
            builder.ListFormat.ListIndent();
        }

        // After the loop, outdent back to the original level.
        for (int i = 0; i < 5; i++)
        {
            builder.ListFormat.ListOutdent();
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ListIndentExample.docx");
        doc.Save(outputPath);
    }
}
