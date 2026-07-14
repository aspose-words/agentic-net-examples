using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three chapters. Each chapter starts with a bookmark named "ChapterStartX".
        for (int i = 1; i <= 3; i++)
        {
            // Define a unique bookmark name for each chapter.
            string bookmarkName = $"ChapterStart{i}";

            // Mark the start of the bookmark at the current cursor position.
            builder.StartBookmark(bookmarkName);

            // Insert a heading that represents the chapter title.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // End the bookmark after the heading.
            builder.EndBookmark(bookmarkName);

            // Add some sample content for the chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}.");
        }

        // Save the document to a file.
        doc.Save("ChapterBookmarks.docx");
    }
}
