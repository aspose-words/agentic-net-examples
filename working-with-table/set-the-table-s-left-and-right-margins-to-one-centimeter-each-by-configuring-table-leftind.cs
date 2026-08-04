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

            // Start a table and add a single cell with some text.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Sample cell content.");
            builder.EndTable();

            // Set the left indent of the table to 1 cm (1 cm = 28.35 points).
            table.LeftIndent = 28.35;

            // Aspose.Words does not provide a RightIndent property.
            // As an alternative, set the right padding of the table to 1 cm.
            // This adds space between the cell contents and the right border,
            // effectively creating a right margin effect.
            table.RightPadding = 28.35;

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableMargins.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
