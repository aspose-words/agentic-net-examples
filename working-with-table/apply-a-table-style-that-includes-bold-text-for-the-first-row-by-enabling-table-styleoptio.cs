using System;
using System.IO;
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

            // Build a simple 2x2 table.
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

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Enable the style for the first row only.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Retrieve the style object and make the first‑row text bold.
            TableStyle tableStyle = (TableStyle)doc.Styles[StyleIdentifier.MediumShading1Accent1];
            tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableStyleFirstRowBold.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
