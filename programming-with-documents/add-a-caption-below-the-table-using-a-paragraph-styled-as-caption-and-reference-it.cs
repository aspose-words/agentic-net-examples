using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCaptionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple 2x2 table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Insert a caption paragraph below the table.
            // Start a bookmark that will be referenced later.
            builder.StartBookmark("Table1");
            // Apply the built‑in Caption style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
            builder.Writeln("Table 1: Sample Table");
            builder.EndBookmark("Table1");

            // Insert a paragraph that references the caption.
            builder.Writeln();
            builder.Write("See ");
            // Insert a REF field that points to the bookmark "Table1".
            builder.InsertField("REF Table1 \\h");
            builder.Writeln(" for details.");

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Optional: Verify that the file was saved.
            if (System.IO.File.Exists(outputPath))
            {
                // Reload the document to ensure it can be opened.
                Document loadedDoc = new Document(outputPath);
                // No further actions needed; the program will exit.
            }
        }
    }
}
