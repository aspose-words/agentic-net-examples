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

            // Add three paragraphs to the document.
            builder.Writeln("Paragraph 1");
            builder.Writeln("Paragraph 2"); // This is the paragraph after which we will insert the table.
            builder.Writeln("Paragraph 3");

            // Locate the paragraph with the exact text "Paragraph 2".
            Paragraph targetParagraph = null;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.GetText().Trim() == "Paragraph 2")
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph == null)
                throw new InvalidOperationException("Target paragraph not found.");

            // Build a simple 2x2 table.
            Table table = new Table(doc);
            // Ensure the table has at least one row.
            table.EnsureMinimum();

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
            // The paragraph's parent is a Body node; use it to perform InsertAfter.
            targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

            // Save the document to a file.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputTableAfterParagraph.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new IOException("Failed to save the output document.");

            // The program ends here; no user interaction required.
        }
    }
}
