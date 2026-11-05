using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph.
            builder.Writeln("This is a sample paragraph.");

            // Add a bulleted list.
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("First bullet item");
            builder.Writeln("Second bullet item");
            builder.Writeln("Third bullet item");
            builder.ListFormat.RemoveNumbers();

            // Add a table with 3 columns and 3 rows.
            Table table = builder.StartTable();

            // Header row.
            for (int i = 0; i < 3; i++)
            {
                builder.InsertCell();
                builder.Writeln($"Header {i + 1}");
                builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }
            builder.EndRow();

            // Data rows.
            for (int row = 0; row < 2; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Writeln($"Row {row + 1}, Cell {col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Save the document as ODT.
            string outputPath = "output.odt";
            doc.Save(outputPath, SaveFormat.Odt);

            // Confirm the file was saved.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
            }
        }
    }
}
