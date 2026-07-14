using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMarginExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a few cells with sample text.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply custom margins around the entire table.
            // LeftIndent sets the left margin.
            table.LeftIndent = 30.0; // points

            // RightIndent property is not available; use DistanceRight to achieve a right margin effect.
            table.DistanceRight = 30.0; // points

            // Define an output path.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);
            string outputPath = Path.Combine(artifactsDir, "TableWithMargins.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
