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

            // Add sample content with heading paragraphs.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is some introductory text.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Details about section 1.1.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("More introductory text.");

            // Iterate through all paragraphs to find headings.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Check if the paragraph style is a heading (any level).
                if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading2 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading3 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading4 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading5 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading6 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading7 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading8 ||
                    para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading9)
                {
                    // Create a new table to insert after the heading.
                    Table table = CreateSampleTable(doc);

                    // Insert the table after the heading paragraph.
                    para.ParentNode.InsertAfter(table, para);
                }
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }

        // Helper method that creates a 2x2 table with sample text.
        private static Table CreateSampleTable(Document doc)
        {
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

            return table;
        }
    }
}
