using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBottomMarginExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndTable();

            // Set the distance between the table bottom and surrounding text (bottom margin) to 5 points.
            table.DistanceBottom = 5.0;

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBottomMargin.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Optionally, you could load the document again to confirm the property persisted.
            Document loadedDoc = new Document(outputPath);
            Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
            if (Math.Abs(loadedTable.DistanceBottom - 5.0) > 0.001)
                throw new InvalidOperationException("The bottom margin was not set correctly.");
        }
    }
}
