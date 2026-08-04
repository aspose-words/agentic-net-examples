using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;

namespace FootnoteTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some main text and a footnote.
            builder.Write("This is a sample paragraph with a footnote.");
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote initial text.");

            // Move the builder into the footnote's paragraph so we can add content there.
            builder.MoveTo(footnote.FirstParagraph);

            // Build a 2x2 table inside the footnote.
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

            // Return the cursor to the end of the main document (optional).
            builder.MoveToDocumentEnd();

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FootnoteTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output file was not created.");

            // The program ends automatically.
        }
    }
}
