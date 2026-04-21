using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Build a simple 2x2 table.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Apply the built‑in "Table Grid" style using the style identifier.
            table.StyleIdentifier = StyleIdentifier.TableGrid;

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableGridStyle.docx");
            doc.Save(outputPath);
        }
    }
}
