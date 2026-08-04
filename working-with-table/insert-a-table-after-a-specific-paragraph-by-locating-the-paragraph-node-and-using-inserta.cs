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

            // Add three paragraphs. The second one will be the reference paragraph.
            builder.Writeln("First paragraph.");
            builder.Writeln("Target paragraph."); // Table will be inserted after this paragraph.
            builder.Writeln("Third paragraph.");

            // Locate the paragraph that contains the exact text "Target paragraph."
            Paragraph targetParagraph = null;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.GetText().Trim() == "Target paragraph.")
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph == null)
                throw new InvalidOperationException("Target paragraph not found.");

            // Create a new table (2 rows x 2 columns) manually.
            Table table = new Table(doc);
            // Ensure the table has at least one row and one cell.
            table.EnsureMinimum();

            // Fill first cell (row 0, column 0).
            table.FirstRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "Cell 1,1"));

            // Add second cell to the first row.
            Cell cell12 = new Cell(doc);
            cell12.AppendChild(new Paragraph(doc));
            cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 1,2"));
            table.FirstRow.AppendChild(cell12);

            // Add second row.
            Row row2 = new Row(doc);
            table.AppendChild(row2);

            // First cell of second row.
            Cell cell21 = new Cell(doc);
            cell21.AppendChild(new Paragraph(doc));
            cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 2,1"));
            row2.AppendChild(cell21);

            // Second cell of second row.
            Cell cell22 = new Cell(doc);
            cell22.AppendChild(new Paragraph(doc));
            cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 2,2"));
            row2.AppendChild(cell22);

            // Insert the table after the target paragraph.
            // The parent of the paragraph is the Body node; use it to perform InsertAfter.
            targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTableAfterParagraph.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new IOException("Failed to save the output document.");

            // The program ends automatically.
        }
    }
}
