using System;
using System.IO;
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

            // Add some paragraphs. One of them contains the keyword we will search for.
            builder.Writeln("This is the first paragraph.");
            builder.Writeln("This paragraph contains the KEYWORD and will be the anchor point.");
            builder.Writeln("This is the third paragraph.");

            // Define the keyword to look for.
            const string keyword = "KEYWORD";

            // Search for the paragraph that contains the keyword.
            Paragraph targetParagraph = null;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.GetText().Contains(keyword))
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph == null)
                throw new InvalidOperationException("Keyword not found in any paragraph.");

            // Build a simple 2x2 table.
            Table table = new Table(doc);

            // First row.
            Row row1 = new Row(doc);
            table.AppendChild(row1);

            Cell cell11 = new Cell(doc);
            cell11.AppendChild(new Paragraph(doc));
            cell11.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 1"));
            row1.AppendChild(cell11);

            Cell cell12 = new Cell(doc);
            cell12.AppendChild(new Paragraph(doc));
            cell12.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 2"));
            row1.AppendChild(cell12);

            // Second row.
            Row row2 = new Row(doc);
            table.AppendChild(row2);

            Cell cell21 = new Cell(doc);
            cell21.AppendChild(new Paragraph(doc));
            cell21.FirstParagraph.AppendChild(new Run(doc, "Row 2, Cell 1"));
            row2.AppendChild(cell21);

            Cell cell22 = new Cell(doc);
            cell22.AppendChild(new Paragraph(doc));
            cell22.FirstParagraph.AppendChild(new Run(doc, "Row 2, Cell 2"));
            row2.AppendChild(cell22);

            // Insert the table after the paragraph that contains the keyword.
            // InsertAfter is a method of the parent node (Body), not of Paragraph itself.
            targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputTableAfterParagraph.docx");
            doc.Save(outputPath);
        }
    }
}
