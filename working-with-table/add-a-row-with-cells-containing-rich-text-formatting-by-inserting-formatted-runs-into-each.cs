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
            builder.StartTable();

            // First row – simple header cells.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row – cells with rich text formatting using formatted runs.
            // First cell: bold text followed by normal text.
            builder.InsertCell();
            builder.Font.Bold = true;
            builder.Write("Bold");
            builder.Font.Bold = false;
            builder.Write(" text");

            // Second cell: italic text followed by normal text.
            builder.InsertCell();
            builder.Font.Italic = true;
            builder.Write("Italic");
            builder.Font.Italic = false;
            builder.Write(" text");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The document was not saved correctly.");
        }
    }
}
