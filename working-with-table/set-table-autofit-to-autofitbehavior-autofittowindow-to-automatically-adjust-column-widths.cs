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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            builder.StartTable();

            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Cell 3");
            builder.InsertCell();
            builder.Writeln("Cell 4");
            builder.EndRow();

            builder.EndTable();

            // Retrieve the created table (first table in the document).
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Apply AutoFit behavior so the table fits the window margins.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            // Save the document to a file.
            string outputPath = "TableAutoFit.docx";
            doc.Save(outputPath);
        }
    }
}
