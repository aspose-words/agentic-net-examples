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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // First cell – set text direction to vertical for East Asian characters.
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
            builder.Write("縦書きテキスト"); // Sample Japanese text.

            // Second cell – normal horizontal text.
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Horizontal;
            builder.Write("Horizontal text");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithVerticalText.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
