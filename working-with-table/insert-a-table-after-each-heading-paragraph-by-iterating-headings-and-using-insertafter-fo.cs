using System;
using System.Collections.Generic;
using System.IO;
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
            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Heading 3");

            // Add a normal paragraph to show that only headings are processed.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is a normal paragraph.");

            // Collect all heading paragraphs first to avoid modifying the collection while iterating.
            List<Paragraph> headingParagraphs = new List<Paragraph>();
            NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in allParagraphs)
            {
                StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
                if (styleId == StyleIdentifier.Heading1 ||
                    styleId == StyleIdentifier.Heading2 ||
                    styleId == StyleIdentifier.Heading3 ||
                    styleId == StyleIdentifier.Heading4 ||
                    styleId == StyleIdentifier.Heading5 ||
                    styleId == StyleIdentifier.Heading6 ||
                    styleId == StyleIdentifier.Heading7 ||
                    styleId == StyleIdentifier.Heading8 ||
                    styleId == StyleIdentifier.Heading9)
                {
                    headingParagraphs.Add(para);
                }
            }

            // For each heading, create a simple 2x2 table and insert it after the heading.
            foreach (Paragraph heading in headingParagraphs)
            {
                Table table = new Table(doc);

                // Build a 2x2 table with sample text.
                for (int rowIdx = 0; rowIdx < 2; rowIdx++)
                {
                    Row row = new Row(doc);
                    table.AppendChild(row);

                    for (int colIdx = 0; colIdx < 2; colIdx++)
                    {
                        Cell cell = new Cell(doc);
                        // Each cell must contain at least one paragraph.
                        Paragraph cellParagraph = new Paragraph(doc);
                        cellParagraph.AppendChild(new Run(doc, $"R{rowIdx + 1}C{colIdx + 1}"));
                        cell.AppendChild(cellParagraph);
                        row.AppendChild(cell);
                    }
                }

                // Insert the table immediately after the heading paragraph.
                // The heading's parent node is typically a Body, which can accept a Table node.
                heading.ParentNode.InsertAfter(table, heading);
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
