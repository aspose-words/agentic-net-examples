using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAutoFitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity (kg)");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Third row.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("40");
            builder.EndRow();

            // Fourth row.
            builder.InsertCell();
            builder.Write("Carrots");
            builder.InsertCell();
            builder.Write("50");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Enable automatic column resizing to fit the cell contents.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document to a file in the current directory.
            string outputPath = "TableAutoFit.docx";
            doc.Save(outputPath);

            // Optional: verify that the file was created.
            if (System.IO.File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to '{outputPath}'.");
            }
            else
            {
                throw new Exception($"Failed to save the document to '{outputPath}'.");
            }
        }
    }
}
