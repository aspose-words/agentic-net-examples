using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace TableVerticalAlignmentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Insert a couple of cells with sample text.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Set the vertical alignment of the table on the page to Center.
            // For floating tables the alignment is controlled by RelativeVerticalAlignment.
            table.RelativeVerticalAlignment = VerticalAlignment.Center;

            // End the table.
            builder.EndTable();

            // Verify that the alignment was applied.
            if (table.RelativeVerticalAlignment != VerticalAlignment.Center)
                throw new InvalidOperationException("Table vertical alignment was not set to Center.");

            // Save the document.
            doc.Save("TableVerticalAlignment.docx");
        }
    }
}
