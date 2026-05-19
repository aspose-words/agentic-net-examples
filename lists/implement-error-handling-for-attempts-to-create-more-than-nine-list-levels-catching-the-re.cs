using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListLevelExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a numbered list (default template has 9 levels: 0‑8).
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

            // Try to assign list levels from 0 to 9.
            // Levels 0‑8 are valid; level 9 exceeds the maximum and will throw.
            for (int i = 0; i <= 9; i++)
            {
                try
                {
                    builder.ListFormat.ListLevelNumber = i; // May throw if i > 8
                    builder.Writeln($"Level {i}");
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    Console.WriteLine($"Caught ArgumentOutOfRangeException for level {i}: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Caught unexpected exception for level {i}: {ex.Message}");
                }
            }

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ListLevelExample.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
