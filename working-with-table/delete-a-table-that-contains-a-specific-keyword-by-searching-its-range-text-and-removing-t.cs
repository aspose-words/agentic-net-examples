using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableDeletion
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the first table (does NOT contain the keyword).
            builder.StartTable();
            builder.InsertCell();
            builder.Write("First table - keep this.");
            builder.EndRow();
            builder.EndTable();

            // Build the second table (contains the keyword "DeleteMe").
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Second table - DeleteMe should be removed.");
            builder.EndRow();
            builder.EndTable();

            // Optional: save the original document for reference.
            string inputPath = Path.Combine(Environment.CurrentDirectory, "Input.docx");
            doc.Save(inputPath);

            // Iterate over all tables in the document and delete those that contain the keyword.
            const string keyword = "DeleteMe";
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

            // Iterate backwards to safely remove nodes while iterating.
            for (int i = tables.Count - 1; i >= 0; i--)
            {
                Table table = (Table)tables[i];
                if (table.Range.Text.Contains(keyword))
                {
                    // Remove the entire table node from the document.
                    table.Remove();
                }
            }

            // Save the modified document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
            doc.Save(outputPath);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");

            // The program finishes without requiring any user interaction.
        }
    }
}
