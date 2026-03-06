using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsLinqDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add sample content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Apple");
            builder.Writeln("Banana");
            builder.Writeln("Cherry");
            builder.Writeln("Date");
            builder.Writeln("Elderberry");

            // -----------------------------------------------------------------
            // LINQ query: select all paragraphs that contain the letter 'a' (case‑insensitive)
            // and order them by the length of their text.
            // -----------------------------------------------------------------
            var paragraphsWithA = doc
                .GetChildNodes(NodeType.Paragraph, true)               // Get all paragraph nodes in the document.
                .OfType<Paragraph>()                                   // Cast to Paragraph type.
                .Where(p => p.GetText().IndexOf('a', StringComparison.OrdinalIgnoreCase) >= 0) // Filter.
                .OrderBy(p => p.GetText().Length)                      // Order by text length.
                .ToList();                                             // Materialize the query.

            // Output the selected paragraphs to the console.
            Console.WriteLine("Paragraphs containing the letter 'a' ordered by length:");
            foreach (var para in paragraphsWithA)
            {
                // Trim removes the trailing paragraph break characters.
                Console.WriteLine($"- \"{para.GetText().Trim()}\" (Length: {para.GetText().Length})");
            }

            // -----------------------------------------------------------------
            // LINQ query: project the paragraph texts into an anonymous type that
            // also contains the word count of each paragraph.
            // -----------------------------------------------------------------
            var paragraphInfo = doc
                .GetChildNodes(NodeType.Paragraph, true)
                .OfType<Paragraph>()
                .Select(p => new
                {
                    Text = p.GetText().Trim(),
                    WordCount = p.GetText()
                                   .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Length
                })
                .ToArray();

            Console.WriteLine("\nParagraphs with word counts:");
            foreach (var info in paragraphInfo)
            {
                Console.WriteLine($"- \"{info.Text}\" – Words: {info.WordCount}");
            }

            // -----------------------------------------------------------------
            // LINQ query: find the first paragraph that starts with the letter 'C'.
            // If found, replace its text with uppercase version.
            // -----------------------------------------------------------------
            var firstCParagraph = doc
                .GetChildNodes(NodeType.Paragraph, true)
                .OfType<Paragraph>()
                .FirstOrDefault(p => p.GetText().StartsWith("C", StringComparison.Ordinal));

            if (firstCParagraph != null)
            {
                // Replace the paragraph's text while preserving the paragraph break.
                string original = firstCParagraph.GetText();
                string withoutBreak = original.TrimEnd('\r', '\n');
                firstCParagraph.Runs.Clear(); // Remove existing runs.
                firstCParagraph.AppendChild(new Run(doc, withoutBreak.ToUpper() + "\r"));
            }

            // Save the modified document.
            doc.Save("LinqDemoOutput.docx");
        }
    }
}
