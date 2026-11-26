using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableInsertionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and add some paragraphs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("First paragraph.");
            builder.Writeln("Target paragraph."); // This paragraph will be the insertion point.
            builder.Writeln("Third paragraph.");

            // Locate the paragraph node that contains the specific text.
            Paragraph targetParagraph = null;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Trim to ignore trailing paragraph marks.
                if (para.GetText().Trim() == "Target paragraph.")
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph == null)
                throw new Exception("Target paragraph not found.");

            // Create a new table.
            Table table = new Table(doc);

            // First row.
            Row row1 = new Row(doc);
            table.Rows.Add(row1);

            Cell cell11 = new Cell(doc);
            cell11.AppendChild(new Paragraph(doc));
            cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
            row1.Cells.Add(cell11);

            Cell cell12 = new Cell(doc);
            cell12.AppendChild(new Paragraph(doc));
            cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
            row1.Cells.Add(cell12);

            // Second row.
            Row row2 = new Row(doc);
            table.Rows.Add(row2);

            Cell cell21 = new Cell(doc);
            cell21.AppendChild(new Paragraph(doc));
            cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
            row2.Cells.Add(cell21);

            Cell cell22 = new Cell(doc);
            cell22.AppendChild(new Paragraph(doc));
            cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
            row2.Cells.Add(cell22);

            // Insert the table after the target paragraph.
            // The parent of the paragraph is a Body node, which can contain a Table.
            targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");

            // The program finishes without waiting for user input.
        }
    }
}
