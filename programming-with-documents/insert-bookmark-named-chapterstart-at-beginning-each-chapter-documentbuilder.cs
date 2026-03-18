using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertChapterBookmarks
{
    static void Main()
    {
        // Load an existing document (replace with your file path) or create a new one.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        int chapterIndex = 1;

        // Iterate through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Identify a chapter heading. Here we treat Heading1 style as a chapter title.
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                // Move the builder's cursor to the start of the heading paragraph.
                builder.MoveTo(para);

                // Create a unique bookmark name for each chapter to avoid duplicate‑name conflicts.
                string bookmarkName = $"ChapterStart{chapterIndex}";

                // Insert a bookmark that starts and ends at the same position (empty bookmark).
                builder.StartBookmark(bookmarkName);
                builder.EndBookmark(bookmarkName);

                chapterIndex++;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
