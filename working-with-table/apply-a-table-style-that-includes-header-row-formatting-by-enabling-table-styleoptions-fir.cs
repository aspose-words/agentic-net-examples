using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Data 1");
            builder.InsertCell();
            builder.Write("Data 2");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Data 3");
            builder.InsertCell();
            builder.Write("Data 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Enable header row formatting via style options.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Auto‑fit the table to its contents.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document.
            const string fileName = "TableWithHeaderStyle.docx";
            doc.Save(fileName);

            // Verify that the file was created.
            if (!File.Exists(fileName))
                throw new InvalidOperationException($"Failed to create {fileName}");

            // Indicate successful completion.
            Console.WriteLine($"Document saved to {Path.GetFullPath(fileName)}");
        }
    }
}
