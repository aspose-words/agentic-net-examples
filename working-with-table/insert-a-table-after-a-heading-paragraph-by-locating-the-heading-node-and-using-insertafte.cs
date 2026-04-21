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
            builder.Writeln("Document introduction.");

            // Insert a heading paragraph (style Heading1) that will be the anchor point.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Target Heading");
            // Reset paragraph formatting for subsequent text.
            builder.ParagraphFormat.ClearFormatting();

            // Add a paragraph after the heading to demonstrate normal flow.
            builder.Writeln("Paragraph following the heading.");

            // Locate the heading paragraph node in the document.
            Paragraph headingParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .FirstOrDefault(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1);

            if (headingParagraph == null)
                throw new InvalidOperationException("Heading paragraph not found.");

            // Build a simple 2x2 table using the Table, Row, and Cell classes.
            Table table = new Table(doc);

            // First row.
            Row firstRow = new Row(doc);
            table.AppendChild(firstRow);

            Cell cell11 = new Cell(doc);
            cell11.AppendChild(new Paragraph(doc));
            cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1,1"));
            firstRow.AppendChild(cell11);

            Cell cell12 = new Cell(doc);
            cell12.AppendChild(new Paragraph(doc));
            cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 1,2"));
            firstRow.AppendChild(cell12);

            // Second row.
            Row secondRow = new Row(doc);
            table.AppendChild(secondRow);

            Cell cell21 = new Cell(doc);
            cell21.AppendChild(new Paragraph(doc));
            cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 2,1"));
            secondRow.AppendChild(cell21);

            Cell cell22 = new Cell(doc);
            cell22.AppendChild(new Paragraph(doc));
            cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 2,2"));
            secondRow.AppendChild(cell22);

            // Insert the table immediately after the heading paragraph.
            // The heading's parent node is a Body, which can accept a Table node.
            headingParagraph.ParentNode.InsertAfter(table, headingParagraph);

            // Define output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableAfterHeading.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved correctly.", outputPath);
        }
    }
}
