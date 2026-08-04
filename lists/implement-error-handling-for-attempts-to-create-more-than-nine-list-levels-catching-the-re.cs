using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Try to set list levels from 0 up to 9 (the valid range is 0‑8).
        // The attempt to set level 9 will throw an exception.
        try
        {
            for (int i = 0; i <= 9; i++)
            {
                // This property throws if the value is outside the 0‑8 range.
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln($"Level {i}");
            }
        }
        catch (Exception ex)
        {
            // Catch the exception and write its message to the console.
            Console.WriteLine("Exception caught while setting list level:");
            Console.WriteLine(ex.Message);
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
