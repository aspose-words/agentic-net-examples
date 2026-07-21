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
            builder.Writeln("Paragraph with the keyword: INSERT_TABLE_HERE.");
            builder.Writeln("This is the third paragraph.");

            // Search for the paragraph that contains the specific keyword.
            const string keyword = "INSERT_TABLE_HERE";
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
                throw new InvalidOperationException($"Keyword \"{keyword}\" not found in the document.");

            // Create a new table that will be inserted after the found paragraph.
            Table table = new Table(doc);
            // Ensure the table has at least one row and one cell.
            table.EnsureMinimum();

            // Populate the table with a single cell containing some text.
            Row row = table.FirstRow;
            Cell cell = row.FirstCell;
            cell.FirstParagraph.AppendChild(new Run(doc, "This is the inserted table cell."));

            // Insert the table after the target paragraph.
            // The paragraph's parent is a Body node, which can accept block-level nodes like Table.
            targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertedTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved correctly.", outputPath);
        }
    }
}
