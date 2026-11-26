using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define number of chapters.
        int chapterCount = 3;

        for (int i = 1; i <= chapterCount; i++)
        {
            // Insert a bookmark named "ChapterStart" at the beginning of the chapter.
            builder.StartBookmark("ChapterStart");

            // Insert a heading for the chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // End the bookmark.
            builder.EndBookmark("ChapterStart");

            // Add some sample content after the heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}.");
            builder.Writeln();
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
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
