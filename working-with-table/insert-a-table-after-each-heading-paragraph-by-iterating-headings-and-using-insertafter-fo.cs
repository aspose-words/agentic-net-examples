using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertion
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample heading paragraphs.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Heading 3");

            // Add a normal paragraph to demonstrate that only headings get tables.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is a normal paragraph without a table.");

            // Iterate through all paragraphs to find headings.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Check if the paragraph style is a heading (Heading1‑Heading9).
                StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
                if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
                {
                    // Create a simple 1‑row, 1‑cell table.
                    Table table = new Table(doc);
                    Row row = new Row(doc);
                    table.AppendChild(row);
                    Cell cell = new Cell(doc);
                    cell.AppendChild(new Paragraph(doc));
                    cell.FirstParagraph.AppendChild(new Run(doc, $"Table after \"{para.GetText().Trim()}\""));
                    row.AppendChild(cell);

                    // Insert the table immediately after the heading paragraph.
                    // The parent of the paragraph is typically a Body node.
                    para.ParentNode.InsertAfter(table, para);
                }
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
