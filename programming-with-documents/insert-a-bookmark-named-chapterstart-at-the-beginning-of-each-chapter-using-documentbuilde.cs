using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three chapters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            // Insert chapter heading (Heading 1 style).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Insert some dummy content for the chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}. It contains several sentences to illustrate the chapter body.");
        }

        // Save the initial document.
        string initialPath = "Chapters.docx";
        doc.Save(initialPath);

        // Load the document to insert bookmarks at the beginning of each chapter.
        Document loadedDoc = new Document(initialPath);
        DocumentBuilder bookmarkBuilder = new DocumentBuilder(loadedDoc);

        // Iterate through all paragraphs and add a bookmark before each Heading 1 paragraph.
        foreach (Paragraph paragraph in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (paragraph.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                // Move the builder to the start of the heading paragraph.
                bookmarkBuilder.MoveTo(paragraph);
                // Insert a bookmark named "ChapterStart".
                bookmarkBuilder.StartBookmark("ChapterStart");
                bookmarkBuilder.EndBookmark("ChapterStart");
            }
        }

        // Save the document with bookmarks.
        string outputPath = "ChaptersWithBookmarks.docx";
        loadedDoc.Save(outputPath);

        // Verify that the output file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
