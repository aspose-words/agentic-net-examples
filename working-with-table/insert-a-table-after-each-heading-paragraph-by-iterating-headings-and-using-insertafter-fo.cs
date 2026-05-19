using System;
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

            // Build sample content with headings.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Some normal paragraph under heading 1.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Another normal paragraph.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Heading 3");

            // Iterate through all paragraphs and insert a table after each heading.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Check if the paragraph style is a heading (Heading1‑Heading9).
                StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
                if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
                {
                    // Create a simple 2×2 table.
                    Table table = CreateSampleTable(doc);

                    // Insert the table after the heading paragraph.
                    // The parent node (usually Body) is a CompositeNode, so cast it to access InsertAfter.
                    ((CompositeNode)para.ParentNode).InsertAfter(table, para);
                }
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }

        // Helper method that builds a 2×2 table with sample text.
        private static Table CreateSampleTable(Document doc)
        {
            Table table = new Table(doc);

            // First row.
            Row row1 = new Row(doc);
            table.AppendChild(row1);

            Cell cell11 = new Cell(doc);
            cell11.AppendChild(new Paragraph(doc));
            cell11.FirstParagraph.AppendChild(new Run(doc, "R1C1"));
            row1.AppendChild(cell11);

            Cell cell12 = new Cell(doc);
            cell12.AppendChild(new Paragraph(doc));
            cell12.FirstParagraph.AppendChild(new Run(doc, "R1C2"));
            row1.AppendChild(cell12);

            // Second row.
            Row row2 = new Row(doc);
            table.AppendChild(row2);

            Cell cell21 = new Cell(doc);
            cell21.AppendChild(new Paragraph(doc));
            cell21.FirstParagraph.AppendChild(new Run(doc, "R2C1"));
            row2.AppendChild(cell21);

            Cell cell22 = new Cell(doc);
            cell22.AppendChild(new Paragraph(doc));
            cell22.FirstParagraph.AppendChild(new Run(doc, "R2C2"));
            row2.AppendChild(cell22);

            return table;
        }
    }
}
