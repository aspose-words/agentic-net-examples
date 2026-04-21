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

            // Insert the first cell.
            builder.InsertCell();

            // Set the text orientation to vertical for East Asian characters.
            // TextOrientation.VerticalFarEast makes Far East characters appear vertically.
            builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;

            // Write some East Asian text.
            builder.Write("こんにちは"); // Japanese greeting

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerticalTextTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Inform that the process completed.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
