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

            // Build the first table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1, Cell 1");
            builder.InsertCell();
            builder.Write("Table 1, Cell 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Table 1, Cell 3");
            builder.InsertCell();
            builder.Write("Table 1, Cell 4");
            builder.EndTable();

            // Insert an empty paragraph to separate the tables.
            builder.InsertParagraph();

            // Build the second table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2, Cell 1");
            builder.InsertCell();
            builder.Write("Table 2, Cell 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Table 2, Cell 3");
            builder.InsertCell();
            builder.Write("Table 2, Cell 4");
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
