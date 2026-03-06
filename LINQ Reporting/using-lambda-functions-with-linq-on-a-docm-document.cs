using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsLinqExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCM document from the file system.
            // This uses the Document constructor (create/load rule) which automatically detects the format.
            Document doc = new Document("InputDocument.docm");

            // Retrieve all Run nodes in the document (including those inside tables, headers, footers, etc.).
            // The Cast<Run>() converts the NodeCollection to an IEnumerable<Run> so we can apply LINQ.
            var runs = doc.GetChildNodes(NodeType.Run, true)
                          .Cast<Run>()
                          // Filter runs that contain the placeholder text "{{Name}}".
                          .Where(r => r.Text.Contains("{{Name}}"));

            // Replace the placeholder with an actual value.
            foreach (Run run in runs)
            {
                run.Text = run.Text.Replace("{{Name}}", "John Doe");
            }

            // Example of using LINQ to count how many paragraphs contain the word "Important".
            int importantParagraphCount = doc.GetChildNodes(NodeType.Paragraph, true)
                                             .Cast<Paragraph>()
                                             .Count(p => p.GetText().Contains("Important"));

            Console.WriteLine($"Number of paragraphs containing 'Important': {importantParagraphCount}");

            // Save the modified document to a new file.
            // This uses the Document.Save(string) method (save rule) which determines the format from the extension.
            doc.Save("OutputDocument.docx");
        }
    }
}
