using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableDeletion
{
    public class TableRemover
    {
        /// <summary>
        /// Deletes every table that contains the specified keyword.
        /// </summary>
        /// <param name="inputPath">Path to the source document.</param>
        /// <param name="outputPath">Path where the modified document will be saved.</param>
        /// <param name="keyword">Keyword to search for inside tables.</param>
        public static void DeleteTableWithKeyword(string inputPath, string outputPath, string keyword)
        {
            // Load the document from disk.
            Document doc = new Document(inputPath);

            // Retrieve all tables in the document (including nested ones).
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

            // Iterate backwards so that removing a table does not affect the loop index.
            for (int i = tables.Count - 1; i >= 0; i--)
            {
                Table table = (Table)tables[i];

                // If the table's text contains the keyword, remove the whole table node.
                if (table.Range.Text?.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    table.Remove();
                }
            }

            // Save the modified document.
            doc.Save(outputPath);
        }

        // Example usage.
        public static void Main()
        {
            string sourceFile = @"C:\Docs\Input.docx";
            string destinationFile = @"C:\Docs\Output.docx";
            string searchKeyword = "CONFIDENTIAL";

            DeleteTableWithKeyword(sourceFile, destinationFile, searchKeyword);
            Console.WriteLine("Tables containing the keyword have been removed.");
        }
    }
}
