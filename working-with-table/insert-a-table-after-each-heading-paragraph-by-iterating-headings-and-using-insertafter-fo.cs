using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace InsertTableAfterHeadings
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample heading paragraphs and some normal text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Introductory paragraph.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Details about the section.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Another introductory paragraph.");

            // Collect all heading paragraphs (Heading1‑Heading9).
            List<Paragraph> headingParagraphs = new List<Paragraph>();
            NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in allParagraphs)
            {
                StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
                if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
                {
                    headingParagraphs.Add(para);
                }
            }

            // Insert a table after each heading using InsertAfter.
            foreach (Paragraph heading in headingParagraphs)
            {
                Table table = CreateSampleTable(doc);
                // Insert the table node directly after the heading paragraph.
                heading.ParentNode.InsertAfter(table, heading);
            }

            // Save the resulting document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }

        // Helper method that creates a simple 2x2 table with sample text.
        private static Table CreateSampleTable(Document doc)
        {
            Table table = new Table(doc);
            // Apply a simple border to make the table visible.
            table.SetBorders(LineStyle.Single, 1.0, Color.Black);

            for (int rowIdx = 0; rowIdx < 2; rowIdx++)
            {
                Row row = new Row(doc);
                table.Rows.Add(row);
                for (int colIdx = 0; colIdx < 2; colIdx++)
                {
                    Cell cell = new Cell(doc);
                    row.Cells.Add(cell);

                    Paragraph para = new Paragraph(doc);
                    Run run = new Run(doc, $"R{rowIdx + 1}C{colIdx + 1}");
                    para.AppendChild(run);
                    cell.AppendChild(para);
                }
            }

            return table;
        }
    }
}
