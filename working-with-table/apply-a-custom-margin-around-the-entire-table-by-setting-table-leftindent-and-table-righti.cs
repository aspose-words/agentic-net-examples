using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMarginExample
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

            // Insert first row with two cells.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Insert second row with two cells.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Apply custom left indent (margin) to the table.
            table.LeftIndent = 30; // points

            // Apply custom right margin using DistanceRight (RightIndent is not allowed).
            table.DistanceRight = 30; // points

            // Save the document.
            string outputPath = "TableMargin.docx";
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!System.IO.File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");
        }
    }
}
