using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder attached to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph and apply a built‑in style (Heading1).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Styled paragraph – the table will be inserted after this line.");

            // Retrieve the paragraph we have just added.
            Paragraph styledParagraph = doc.FirstSection.Body.LastParagraph;

            // Move the builder's cursor to the styled paragraph.
            builder.MoveTo(styledParagraph);

            // Insert an empty paragraph after the styled paragraph to position the cursor
            // where the table should be placed.
            builder.InsertParagraph();

            // Build a 2×2 table at the current cursor position.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableInStyledParagraph.docx");
            doc.Save(outputPath);

            // Simple validation – ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // The program ends automatically; no user interaction is required.
        }
    }
}
