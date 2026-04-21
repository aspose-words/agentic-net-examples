using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMarginsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 1x1 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Sample cell");
            builder.EndTable();

            // Set left indent to 1 centimeter (1 cm = 28.35 points).
            table.LeftIndent = 28.35;

            // Set right distance (margin) to 1 centimeter.
            table.DistanceRight = 28.35;

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableMargins.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");

            // Optionally, you could open the document to verify the settings,
            // but this example ends after saving.
        }
    }
}
