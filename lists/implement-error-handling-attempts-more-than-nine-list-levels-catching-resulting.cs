using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ListLevelErrorHandling
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a default numbered list to the builder.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Attempt to set list levels from 0 up to 9 (inclusive).
        // Valid levels are 0‑8; level 9 should raise an exception.
        for (int i = 0; i <= 9; i++)
        {
            try
            {
                // This will succeed for i = 0..8 and fail for i = 9.
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln($"Level {i}");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Catch the specific exception thrown when the level is out of range.
                Console.WriteLine($"Error: Cannot set ListLevelNumber to {i}. {ex.Message}");
                // Optionally, break out of the loop if further processing is not needed.
                break;
            }
            catch (Exception ex)
            {
                // Fallback for any other unexpected exceptions.
                Console.WriteLine($"Unexpected error at level {i}: {ex.Message}");
                break;
            }
        }

        // Reset the List property to stop list formatting.
        builder.ListFormat.List = null;

        // Save the document.
        doc.Save("Lists_LevelErrorHandling.docx");
    }
}
