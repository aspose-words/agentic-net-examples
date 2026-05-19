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

            // Add three paragraphs to the document.
            builder.Writeln("Paragraph 1");
            builder.Writeln("Paragraph 2");
            builder.Writeln("Paragraph 3");

            // Locate the paragraph that contains the text "Paragraph 2".
            Paragraph targetParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .First(p => p.GetText().Trim() == "Paragraph 2");

            // Build a simple 2x2 table using the Table, Row, and Cell classes.
            Table table = new Table(doc);

            // First row.
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

            // Second row.
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

            // Insert the table after the located paragraph.
            targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAfterParagraph.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
