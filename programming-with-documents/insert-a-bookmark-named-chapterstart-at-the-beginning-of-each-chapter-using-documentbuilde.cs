using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace BookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert three sample chapters.
            for (int i = 1; i <= 3; i++)
            {
                // Insert an empty bookmark named "ChapterStartX" at the current cursor position.
                // The bookmark is placed before the chapter heading, i.e., at the beginning of the chapter.
                string bookmarkName = $"ChapterStart{i}";
                builder.StartBookmark(bookmarkName);
                builder.EndBookmark(bookmarkName);

                // Write the chapter heading (styled as Heading 1).
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Chapter {i}");

                // Write some dummy paragraph text for the chapter body.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln($"This is the content of chapter {i}. It demonstrates how to place a bookmark at the start of each chapter.");
                builder.Writeln(); // Add an empty line between chapters.
            }

            // Save the document to the local file system.
            string outputPath = "ChapterBookmarks.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
