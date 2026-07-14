using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph before the table.
            builder.Writeln("Paragraph before the original table.");

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            // First row.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();
            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();
            // Finish the table.
            builder.EndTable();

            // Insert another paragraph after the table.
            builder.Writeln("Paragraph after the original table.");

            // Locate the first paragraph (the one before the table).
            Paragraph referenceParagraph = doc.FirstSection.Body.FirstParagraph;

            // Locate the original table node.
            Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Clone the table (deep clone).
            Table clonedTable = (Table)originalTable.Clone(true);

            // Insert the cloned table after the reference paragraph.
            // The parent of the paragraph is the Body node.
            Body body = doc.FirstSection.Body;
            body.InsertAfter(clonedTable, referenceParagraph);

            // Save the document to a local file.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableCloneExample.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
