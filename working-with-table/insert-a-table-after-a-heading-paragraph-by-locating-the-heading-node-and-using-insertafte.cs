using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some content before the heading.
            builder.Writeln("Introduction paragraph.");

            // Insert a heading paragraph (style Heading1).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Sample Heading");
            // Reset style for following paragraphs.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Paragraph after heading.");

            // Locate the heading paragraph we just added.
            Paragraph headingParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .First(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1);

            // Create a new table node.
            Table table = new Table(doc);
            // Ensure the table has at least one row and one cell.
            table.EnsureMinimum();

            // Populate the first row with two cells.
            Cell cell1 = table.FirstRow.FirstCell;
            cell1.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));

            Cell cell2 = new Cell(doc);
            cell2.AppendChild(new Paragraph(doc));
            cell2.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
            table.FirstRow.AppendChild(cell2);

            // Add a second row with two cells.
            Row secondRow = new Row(doc);
            table.AppendChild(secondRow);

            Cell cell3 = new Cell(doc);
            cell3.AppendChild(new Paragraph(doc));
            cell3.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
            secondRow.AppendChild(cell3);

            Cell cell4 = new Cell(doc);
            cell4.AppendChild(new Paragraph(doc));
            cell4.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
            secondRow.AppendChild(cell4);

            // Insert the table after the heading paragraph.
            // The heading's parent is a Body node, which can accept block-level nodes.
            headingParagraph.ParentNode.InsertAfter(table, headingParagraph);

            // Define output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableAfterHeading.docx");
            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // No interactive prompts; program ends here.
        }
    }
}
