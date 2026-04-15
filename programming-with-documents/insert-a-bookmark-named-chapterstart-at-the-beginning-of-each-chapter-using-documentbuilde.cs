using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace BookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build three chapters. Each chapter starts with a bookmark named "ChapterStart_X".
            for (int chapter = 1; chapter <= 3; chapter++)
            {
                // Insert a bookmark at the very beginning of the chapter.
                string bookmarkName = $"ChapterStart_{chapter}";
                builder.StartBookmark(bookmarkName);

                // Write the chapter heading using Heading1 style.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Chapter {chapter}");

                // Close the bookmark immediately after the heading.
                builder.EndBookmark(bookmarkName);

                // Add some sample paragraph text for the chapter.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln($"This is the content of chapter {chapter}. It demonstrates how to insert a bookmark at the start of each chapter.");

                // Insert a page break before the next chapter (except after the last one).
                if (chapter < 3)
                {
                    builder.InsertBreak(BreakType.PageBreak);
                }
            }

            // Save the document to the local file system.
            string outputPath = "ChapterBookmarks.docx";
            doc.Save(outputPath);
        }
    }
}
