using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableIndentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Configure paragraph indentation.
            builder.ParagraphFormat.LeftIndent = 30;   // 30 points left indent.
            builder.ParagraphFormat.RightIndent = 30; // 30 points right indent.

            // Start a table.
            Table table = builder.StartTable();

            // Insert the first cell (required before setting table formatting).
            builder.InsertCell();

            // Align the table's left indent with the paragraph's left indent.
            table.LeftIndent = builder.ParagraphFormat.LeftIndent;

            // Populate the first row with two cells.
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableIndent.docx");
            doc.Save(outputPath);
        }
    }
}
