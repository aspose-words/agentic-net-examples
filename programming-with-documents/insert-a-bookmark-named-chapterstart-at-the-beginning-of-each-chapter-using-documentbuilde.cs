using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample chapters (Heading 1 style) with some content.
        for (int i = 1; i <= 3; i++)
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}. It contains several sentences to illustrate the chapter body.");
        }

        // Save the initial document (optional, shows the intermediate state).
        doc.Save("Sample.docx");

        // Insert a bookmark named "ChapterStartX" at the beginning of each chapter.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        int chapterIndex = 1;

        foreach (Paragraph para in paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                // Move the builder to the start of the heading paragraph.
                DocumentBuilder bmBuilder = new DocumentBuilder(doc);
                bmBuilder.MoveTo(para);

                // Create a unique bookmark name for each chapter.
                string bookmarkName = $"ChapterStart{chapterIndex}";
                bmBuilder.StartBookmark(bookmarkName);
                bmBuilder.EndBookmark(bookmarkName);

                chapterIndex++;
            }
        }

        // Save the final document with bookmarks.
        doc.Save("ChaptersWithBookmarks.docx");
    }
}
