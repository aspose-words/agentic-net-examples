using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build three sample chapters. Each chapter starts with a Heading 1 paragraph.
        for (int i = 1; i <= 3; i++)
        {
            // Chapter title.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Sample body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}. It contains several sentences to illustrate the chapter body.");
            builder.Writeln();
        }

        // Insert a bookmark named "ChapterStart" at the beginning of each chapter.
        // We iterate over all paragraphs, locate those with Heading1 style, move the builder to the start of the paragraph,
        // and add a uniquely named bookmark (ChapterStart_1, ChapterStart_2, ...).
        int chapterIndex = 1;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                // Move the builder's cursor to the start of the heading paragraph.
                builder.MoveTo(para);
                // Insert the bookmark.
                string bookmarkName = $"ChapterStart_{chapterIndex}";
                builder.StartBookmark(bookmarkName);
                // Optionally, you could write something inside the bookmark or just close it immediately.
                builder.EndBookmark(bookmarkName);
                chapterIndex++;
            }
        }

        // Save the document to the local file system.
        doc.Save("Output.docx");
    }
}
