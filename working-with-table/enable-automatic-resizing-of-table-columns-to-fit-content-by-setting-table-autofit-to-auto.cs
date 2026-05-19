using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAutoFitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to construct a simple table.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Item");

            // First row, second cell with longer content.
            builder.InsertCell();
            builder.Write("Description with a considerably longer text that should cause the column to expand.");

            // End the first row.
            builder.EndRow();

            // Add a second row.
            builder.InsertCell();
            builder.Write("Apple");

            builder.InsertCell();
            builder.Write("A fruit that is typically red, green, or yellow.");

            builder.EndRow();

            // Finish building the table.
            builder.EndTable();

            // Enable automatic column resizing to fit the cell contents.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "TableAutoFitToContents.docx");
            doc.Save(outputPath);
        }
    }
}
