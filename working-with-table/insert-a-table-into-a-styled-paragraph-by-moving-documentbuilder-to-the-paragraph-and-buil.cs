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

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a styled paragraph (Heading1) that will contain the table.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Styled Paragraph");

            // Retrieve the paragraph we just added.
            Paragraph styledParagraph = doc.FirstSection.Body.LastParagraph;

            // Move the builder's cursor to the styled paragraph.
            builder.MoveTo(styledParagraph);

            // Start building a table at the current cursor position.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "TableInStyledParagraph.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");

            // The program ends automatically; no user interaction required.
        }
    }
}
