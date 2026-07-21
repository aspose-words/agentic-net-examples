using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTableInHeaderFooter
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder cursor to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Start building a table inside the header.
            Table headerTable = builder.StartTable();

            // First row of the table.
            builder.InsertCell();
            builder.Write("Header Cell 1");
            builder.InsertCell();
            builder.Write("Header Cell 2");
            builder.EndRow();

            // Second row of the table.
            builder.InsertCell();
            builder.Write("Header Cell 3");
            builder.InsertCell();
            builder.Write("Header Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Return the cursor to the main body of the document.
            builder.MoveToSection(0);
            builder.Writeln("Body content starts here.");

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "HeaderFooterTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Indicate successful completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
