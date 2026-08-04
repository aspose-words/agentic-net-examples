using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAllowOverlapExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with one cell.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Floating table cell content.");
            builder.EndTable();

            // The table is an inline table by default, which allows overlap (AllowOverlap == true).
            // No need to set TextWrapping or other floating properties because they would
            // change the AllowOverlap behavior.

            // Verify that overlapping is enabled.
            if (!table.AllowOverlap)
                throw new InvalidOperationException("Table does not allow overlap, but it should.");

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAllowOverlap.docx");
            doc.Save(outputPath);

            // Ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("Failed to save the document.", outputPath);

            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
    }
}
