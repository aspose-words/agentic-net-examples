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

            // Add some paragraphs to the document.
            builder.Writeln("Paragraph 1");
            builder.Writeln("Paragraph 2 - target"); // This paragraph will be the insertion point.
            builder.Writeln("Paragraph 3");

            // Locate the paragraph node that contains the target text.
            Paragraph targetParagraph = null;
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.GetText().Trim() == "Paragraph 2 - target")
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph == null)
                throw new InvalidOperationException("Target paragraph not found.");

            // Build a simple 2x2 table using the Table, Row, and Cell classes.
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

            // Insert the table after the located paragraph.
            // The InsertAfter method is called on the parent node of the reference paragraph.
            targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

            // Save the document to a file.
            string outputPath = "TableInsertedAfterParagraph.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new IOException($"Failed to create the output file: {outputPath}");
        }
    }
}
