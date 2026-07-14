using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // First row (header).
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row (data).
            builder.InsertCell();
            builder.Write("Data 1");
            builder.InsertCell();
            builder.Write("Data 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply a built‑in style to the table.
            table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

            // Disable row banding by clearing all style options.
            // TableStyleOptions.RowBands is the flag that enables banding,
            // so using TableStyleOptions.None removes it.
            table.StyleOptions = TableStyleOptions.None;

            // Save the document.
            doc.Save("TableStyleNoRowBanding.docx");
        }
    }
}
