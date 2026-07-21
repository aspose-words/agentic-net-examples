using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use a DocumentBuilder to add sample chapters (Heading1) and body text.
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            // Insert a chapter heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Insert some paragraph text for the chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}.");
            builder.Writeln();
        }

        // Insert a bookmark named "ChapterStart" at the beginning of each chapter.
        // A chapter is identified by a paragraph with the Heading1 style.
        DocumentBuilder bookmarkBuilder = new DocumentBuilder(doc);
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        foreach (Paragraph para in paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                // Move the builder's cursor to the start of the heading paragraph.
                bookmarkBuilder.MoveTo(para);
                // Create a zero‑length bookmark at this position.
                bookmarkBuilder.StartBookmark("ChapterStart");
                bookmarkBuilder.EndBookmark("ChapterStart");
            }
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChapterBookmarks.docx");
        doc.Save(outputPath);

        // Optional: verify that the file was created.
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
