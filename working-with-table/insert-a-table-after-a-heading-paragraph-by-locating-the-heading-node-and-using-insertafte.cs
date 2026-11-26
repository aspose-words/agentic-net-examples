using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableInsertionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a heading paragraph.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Sample Heading");

            // Locate the heading paragraph we just added.
            Paragraph headingParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .First(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1);

            // Build a simple 2x2 table.
            Table table = new Table(doc);

            // First row.
            Row row1 = new Row(doc);
            Cell cell11 = new Cell(doc);
            cell11.AppendChild(new Paragraph(doc));
            cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
            row1.AppendChild(cell11);

            Cell cell12 = new Cell(doc);
            cell12.AppendChild(new Paragraph(doc));
            cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
            row1.AppendChild(cell12);
            table.AppendChild(row1);

            // Second row.
            Row row2 = new Row(doc);
            Cell cell21 = new Cell(doc);
            cell21.AppendChild(new Paragraph(doc));
            cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
            row2.AppendChild(cell21);

            Cell cell22 = new Cell(doc);
            cell22.AppendChild(new Paragraph(doc));
            cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
            row2.AppendChild(cell22);
            table.AppendChild(row2);

            // Insert the table after the heading paragraph.
            headingParagraph.ParentNode.InsertAfter(table, headingParagraph);

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
