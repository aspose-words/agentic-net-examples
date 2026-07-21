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

            // First cell – set text orientation to vertical for East Asian characters.
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
            builder.Write("垂直文本"); // Sample East Asian text.

            // Second cell – normal horizontal orientation.
            builder.InsertCell();
            builder.Write("Normal");

            // End the first row and the table.
            builder.EndRow();
            builder.EndTable();

            // Define output path and ensure the directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "VerticalTextDirection.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
