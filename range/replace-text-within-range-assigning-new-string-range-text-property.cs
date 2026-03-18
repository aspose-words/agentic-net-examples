using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceTextInRange
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample text that we will replace later.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear _Customer_,");
        builder.Writeln("Thank you for your purchase.");

        // Replace the placeholder "_Customer_" with the actual name.
        // The Range.Replace method works on the whole document range.
        // It returns the number of replacements performed (optional to capture).
        int replacedCount = doc.Range.Replace("_Customer_", "John Doe");

        // Optional: verify that the replacement occurred.
        Console.WriteLine($"Replacements made: {replacedCount}");
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(doc.GetText().Trim());

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
