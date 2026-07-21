using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertion
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample heading paragraphs.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Some introductory text.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Details about section 1.1.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("More content.");

            // Collect all heading paragraphs first to avoid modifying the collection while iterating.
            List<Paragraph> headingParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .Where(p => IsHeading(p))
                .ToList();

            // Insert a table after each heading.
            foreach (Paragraph heading in headingParagraphs)
            {
                Table table = CreateSampleTable(doc);
                // Insert the table after the heading paragraph.
                heading.ParentNode.InsertAfter(table, heading);
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
        }

        // Determines whether a paragraph is a heading (any heading level).
        private static bool IsHeading(Paragraph paragraph)
        {
            StyleIdentifier style = paragraph.ParagraphFormat.StyleIdentifier;
            return style == StyleIdentifier.Heading1 ||
                   style == StyleIdentifier.Heading2 ||
                   style == StyleIdentifier.Heading3 ||
                   style == StyleIdentifier.Heading4 ||
                   style == StyleIdentifier.Heading5 ||
                   style == StyleIdentifier.Heading6 ||
                   style == StyleIdentifier.Heading7 ||
                   style == StyleIdentifier.Heading8 ||
                   style == StyleIdentifier.Heading9;
        }

        // Creates a simple 2x2 table with sample text.
        private static Table CreateSampleTable(Document doc)
        {
            Table table = new Table(doc);

            // First row
            Row row1 = new Row(doc);
            table.AppendChild(row1);
            Cell cell11 = new Cell(doc);
            cell11.AppendChild(new Paragraph(doc));
            cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1,1"));
            row1.AppendChild(cell11);
            Cell cell12 = new Cell(doc);
            cell12.AppendChild(new Paragraph(doc));
            cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 1,2"));
            row1.AppendChild(cell12);

            // Second row
            Row row2 = new Row(doc);
            table.AppendChild(row2);
            Cell cell21 = new Cell(doc);
            cell21.AppendChild(new Paragraph(doc));
            cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 2,1"));
            row2.AppendChild(cell21);
            Cell cell22 = new Cell(doc);
            cell22.AppendChild(new Paragraph(doc));
            cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 2,2"));
            row2.AppendChild(cell22);

            return table;
        }
    }
}
