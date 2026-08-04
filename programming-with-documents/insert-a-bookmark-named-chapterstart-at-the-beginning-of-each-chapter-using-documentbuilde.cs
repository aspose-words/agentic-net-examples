using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the number of chapters we want to create.
        int chapterCount = 3;

        for (int i = 1; i <= chapterCount; i++)
        {
            // Insert a bookmark that marks the start of the current chapter.
            // Use a unique name for each bookmark to avoid duplicate‑name conflicts.
            string bookmarkName = $"ChapterStart_{i}";
            builder.StartBookmark(bookmarkName);

            // Write the chapter heading (styled as Heading 1) and some sample text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}.");

            // Close the bookmark after the chapter content.
            builder.EndBookmark(bookmarkName);
        }

        // Save the document to the local file system.
        string outputPath = "ChapterBookmarks.docx";
        doc.Save(outputPath);
    }
}
