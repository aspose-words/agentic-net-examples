using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListLevelDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a numbered list using the default template.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

            // Try to set list levels from 0 up to 9 (valid range is 0‑8).
            // The attempt to set level 9 should throw an exception.
            for (int level = 0; level <= 9; level++)
            {
                try
                {
                    builder.ListFormat.ListLevelNumber = level; // May throw if level > 8
                    builder.Writeln($"Level {level}");
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    // Catch the specific exception thrown for an invalid list level.
                    Console.WriteLine($"Caught exception when setting level {level}: {ex.Message}");
                    // Stop adding further items after the error.
                    break;
                }
                catch (Exception ex)
                {
                    // Catch any other unexpected exceptions.
                    Console.WriteLine($"Unexpected error at level {level}: {ex.Message}");
                    break;
                }
            }

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the output file.
            doc.Save("ListLevelDemo.docx");
        }
    }
}
