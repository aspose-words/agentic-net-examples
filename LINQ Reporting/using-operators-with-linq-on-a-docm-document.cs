using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsLinqDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCM file (macro-enabled Word document).
            // The file must exist on disk; replace with your actual file path.
            string inputPath = @"C:\Docs\Sample.docm";

            // Load the DOCM document using the Document constructor (lifecycle rule).
            Document doc = new Document(inputPath);

            // -----------------------------------------------------------------
            // Example 1: Get all paragraphs that contain the word "Aspose"
            // and print their text.
            // -----------------------------------------------------------------
            var paragraphsWithAspose = doc
                .GetChildNodes(NodeType.Paragraph, true)               // Get all paragraph nodes (including those in tables, headers, etc.)
                .OfType<Paragraph>()                                   // Cast to Paragraph type for LINQ
                .Where(p => p.GetText().IndexOf("Aspose", StringComparison.OrdinalIgnoreCase) >= 0);

            Console.WriteLine("Paragraphs containing the word \"Aspose\":");
            foreach (var para in paragraphsWithAspose)
                Console.WriteLine(para.GetText().Trim());

            // -----------------------------------------------------------------
            // Example 2: Count the total number of words in the document.
            // -----------------------------------------------------------------
            int totalWordCount = doc
                .GetChildNodes(NodeType.Paragraph, true)
                .OfType<Paragraph>()
                .Select(p => p.GetText())
                .SelectMany(text => text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                .Count();

            Console.WriteLine($"\nTotal word count (approximate): {totalWordCount}");

            // -----------------------------------------------------------------
            // Example 3: List all tables and the number of rows in each table,
            // ordered by row count descending.
            // -----------------------------------------------------------------
            var tableInfo = doc
                .GetChildNodes(NodeType.Table, true)
                .OfType<Table>()
                .Select(t => new
                {
                    Table = t,
                    RowCount = t.Rows.Count
                })
                .OrderByDescending(x => x.RowCount);

            Console.WriteLine("\nTables sorted by row count (descending):");
            foreach (var info in tableInfo)
                Console.WriteLine($"Table with {info.RowCount} rows.");

            // -----------------------------------------------------------------
            // Example 4: Group paragraphs by their style identifier and display the group sizes.
            // -----------------------------------------------------------------
            var paragraphsByStyle = doc
                .GetChildNodes(NodeType.Paragraph, true)
                .OfType<Paragraph>()
                .GroupBy(p => p.ParagraphFormat.StyleIdentifier);

            Console.WriteLine("\nParagraph count per style:");
            foreach (var group in paragraphsByStyle)
            {
                string styleName = group.Key.ToString();
                int count = group.Count();
                Console.WriteLine($"{styleName}: {count} paragraph(s)");
            }

            // -----------------------------------------------------------------
            // Save the (unchanged) document to a new file.
            // The Save method determines the format from the file extension (lifecycle rule).
            // -----------------------------------------------------------------
            string outputPath = @"C:\Docs\Processed.docx";
            doc.Save(outputPath);

            Console.WriteLine($"\nDocument saved to: {outputPath}");
        }
    }
}
