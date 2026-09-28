using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("This is a sample document.");
        builder.Writeln("ReplaceMe"); // Text that will be replaced.

        // Validate that the document's range contains no form fields.
        if (doc.Range.FormFields.Count == 0)
        {
            // Perform the text replacement because no form fields are present.
            FindReplaceOptions options = new FindReplaceOptions();
            doc.Range.Replace("ReplaceMe", "ReplacedText", options);
        }
        else
        {
            // If form fields exist, skip replacement (could log or handle as needed).
            Console.WriteLine("Document contains form fields; replacement aborted.");
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Optional verification: ensure the replacement occurred.
        bool replacementSucceeded = doc.Range.Text.Contains("ReplacedText");
        Console.WriteLine(replacementSucceeded
            ? "Replacement completed successfully."
            : "Replacement was not performed.");
    }
}
