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

            // Add a paragraph that contains the keyword we will search for.
            const string keyword = "INSERT_TABLE_AFTER_ME";
            builder.Writeln($"This paragraph contains the keyword: {keyword}.");
            // Add another paragraph to demonstrate normal content.
            builder.Writeln("This is another paragraph that should stay before the table.");

            // Search for the paragraph node that contains the keyword.
            Paragraph targetParagraph = null;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Get the plain text of the paragraph.
                string text = para.GetText();
                if (text != null && text.Contains(keyword))
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph == null)
                throw new InvalidOperationException("Target paragraph with the keyword was not found.");

            // Build a simple 2x2 table manually (without using DocumentBuilder.StartTable).
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

            // Insert the table after the target paragraph using InsertAfter.
            // The paragraph is a child of the Body node, so we insert after it in the Body.
            Body body = doc.FirstSection.Body;
            body.InsertAfter(table, targetParagraph);

            // Validate that the insertion succeeded.
            if (targetParagraph.NextSibling != table)
                throw new InvalidOperationException("The table was not inserted after the target paragraph.");

            // Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputTableAfterParagraph.docx");
            doc.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
