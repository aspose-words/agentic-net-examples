using System;
using Aspose.Words;

namespace AsposeWordsTocExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents with the switch to include heading levels 1 through 3.
            // \\o "1-3" specifies the entry levels, \\h enables hyperlinks, \\z hides page numbers in web layout,
            // and \\u builds the TOC from outline levels.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Add a page break after the TOC so headings start on a new page.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert Heading 1.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1: Introduction");

            // Insert some body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the introductory chapter.");

            // Insert Heading 2.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1: Background");

            // Insert more body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Background information goes here.");

            // Insert Heading 3.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Subsection 1.1.1: Details");

            // Insert final body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Detailed discussion.");

            // Update fields to populate the TOC.
            doc.UpdateFields();

            // Save the document to disk.
            string outputPath = "TOC_CustomLevels.docx";
            doc.Save(outputPath);
        }
    }
}
