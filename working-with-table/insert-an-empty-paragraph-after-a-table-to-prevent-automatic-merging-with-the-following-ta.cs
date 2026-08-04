using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableParagraphExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- First table ----------
            builder.StartTable();
            // Row 1
            builder.InsertCell();
            builder.Write("Table 1, Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Table 1, Row 1, Cell 2");
            builder.EndRow();
            // Row 2
            builder.InsertCell();
            builder.Write("Table 1, Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Table 1, Row 2, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Insert an empty paragraph to separate the tables.
            // This prevents Word from automatically merging the two tables.
            builder.InsertParagraph();

            // ---------- Second table ----------
            builder.StartTable();
            // Row 1
            builder.InsertCell();
            builder.Write("Table 2, Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Table 2, Row 1, Cell 2");
            builder.EndRow();
            // Row 2
            builder.InsertCell();
            builder.Write("Table 2, Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Table 2, Row 2, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // The program ends automatically; no user interaction required.
        }
    }
}
