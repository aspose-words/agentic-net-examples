using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace HeaderTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder's cursor to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Start a new table in the header.
            Table table = builder.StartTable();

            // First cell.
            builder.InsertCell();
            builder.Write("Header Cell 1");

            // Second cell.
            builder.InsertCell();
            builder.Write("Header Cell 2");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to disk.
            string outputPath = "HeaderTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
