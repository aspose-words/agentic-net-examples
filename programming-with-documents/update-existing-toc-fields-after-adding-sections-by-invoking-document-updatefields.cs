using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace UpdateTocExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents at the beginning of the document.
            // The TOC will include headings of levels 1 through 3.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.Writeln(); // Add a blank line after TOC.

            // Add first heading (Level 1) and some body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1: Introduction");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the introductory chapter.");

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Add second heading (Level 1) and some body text in the new section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2: Details");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This section contains detailed information.");

            // Update all fields in the document, which refreshes the TOC entries.
            doc.UpdateFields();

            // Save the document to disk.
            string outputPath = "UpdatedTocDocument.docx";
            doc.Save(outputPath);
        }
    }
}
