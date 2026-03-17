using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAutoFit
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to construct a simple table.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();

            // Insert a single row with two cells so the table is valid.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Apply AutoFit behavior so the table expands to fit the page margins.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            // Save the document to disk.
            doc.Save("TableAutoFitToWindow.docx");
        }
    }
}
