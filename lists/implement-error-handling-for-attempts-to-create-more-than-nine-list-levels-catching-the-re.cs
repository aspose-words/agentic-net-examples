using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

            // Try to set list levels beyond the allowed range (0‑8).
            // The ListLevelNumber property accepts values 0 to 8 only.
            // Setting a value of 9 should throw an ArgumentOutOfRangeException.
            try
            {
                for (int level = 0; level <= 9; level++)
                {
                    // This will succeed for levels 0‑8 and fail for level 9.
                    builder.ListFormat.ListLevelNumber = level;
                    builder.Writeln($"Level {level}");
                }
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Handle the exception that occurs when trying to exceed the maximum list level.
                Console.WriteLine("Caught exception: " + ex.Message);
            }
            finally
            {
                // End the list formatting.
                builder.ListFormat.RemoveNumbers();
            }

            // Save the document to the output file.
            doc.Save("Lists_ErrorHandling_Output.docx");
        }
    }
}
