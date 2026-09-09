using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs that contain the text we will replace.
        builder.Writeln("This is the first paragraph. ReplaceMe");
        builder.Writeln("This is the second paragraph. ReplaceMe");
        builder.Writeln("This paragraph does not contain the target text.");

        // Define the text to find and the replacement string.
        // The replacement uses the meta‑character &p to insert a paragraph break after each match.
        string findText = "ReplaceMe";
        string replaceText = "ReplaceMe&p";

        // Perform the find‑and‑replace operation.
        int replacedCount = doc.Range.Replace(findText, replaceText, new FindReplaceOptions());

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: output a simple confirmation to the console.
        Console.WriteLine($"Replacements made: {replacedCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
