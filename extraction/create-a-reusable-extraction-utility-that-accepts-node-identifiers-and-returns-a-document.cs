using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ExtractionExample
{
    public class ExtractionUtility
    {
        // Creates a sample source document with two bookmarks.
        private static Document CreateSampleDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Text before the first bookmark.
            builder.Writeln("Intro paragraph.");

            // First bookmark with styled content.
            builder.StartBookmark("First");
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            builder.Writeln("This is the first bookmarked content.");
            builder.EndBookmark("First");

            // Text between bookmarks.
            builder.Writeln("Intermediate paragraph.");

            // Second bookmark with different styling.
            builder.StartBookmark("Second");
            builder.Font.Italic = true;
            builder.Font.Color = System.Drawing.Color.Blue; // Aspose.Words uses System.Drawing for colors.
            builder.Writeln("Second bookmarked content goes here.");
            builder.EndBookmark("Second");

            // Final paragraph.
            builder.Writeln("Conclusion paragraph.");

            return doc;
        }

        // Extracts the content of the specified bookmark into a new Document.
        public static Document ExtractBookmark(string bookmarkName)
        {
            // Create (or load) the source document.
            Document source = CreateSampleDocument();

            // Locate the bookmark.
            Bookmark bookmark = source.Range.Bookmarks[bookmarkName];
            if (bookmark == null)
                throw new ArgumentException($"Bookmark '{bookmarkName}' does not exist in the source document.");

            // Retrieve the text inside the bookmark.
            string extractedText = bookmark.Text;

            // Build a new document that will contain the extracted content.
            Document result = new Document();
            // Remove the default empty section/paragraph.
            result.RemoveAllChildren();

            // Create a new section.
            Section section = new Section(result);
            result.AppendChild(section);

            // Create a body for the section.
            Body body = new Body(result);
            section.AppendChild(body);

            // Create a paragraph and add the extracted text as a run.
            Paragraph paragraph = new Paragraph(result);
            body.AppendChild(paragraph);
            Run run = new Run(result, extractedText);
            paragraph.AppendChild(run);

            return result;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Extract the content of the bookmark named "First".
            Document extracted = ExtractionUtility.ExtractBookmark("First");

            // Save the extracted document to a local file.
            extracted.Save("Extracted_First.docx");

            // Extract the content of the bookmark named "Second".
            Document extractedSecond = ExtractionUtility.ExtractBookmark("Second");
            extractedSecond.Save("Extracted_Second.docx");
        }
    }
}
