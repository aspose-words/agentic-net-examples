using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text with multiple manual line breaks (\v) between words.
        // ControlChar.LineBreak represents a manual line break.
        builder.Write("First line");
        builder.Write(ControlChar.LineBreak);
        builder.Write(ControlChar.LineBreak); // two line breaks in a row
        builder.Write("Second line");
        builder.Write(ControlChar.LineBreak);
        builder.Write(ControlChar.LineBreak);
        builder.Write(ControlChar.LineBreak); // three line breaks in a row
        builder.Write("Third line");
        builder.Writeln(); // end of paragraph

        // Define a regular expression that matches two or more consecutive line break characters.
        // ControlChar.LineBreak is "\v", which is the same as the regex escape \v.
        Regex multipleLineBreaks = new Regex(@"(\v){2,}");

        // Replace each occurrence with a single paragraph break using the "\n" metacharacter.
        // In Aspose.Words, "\n" is interpreted as a line feed; when saving to Word formats it becomes a paragraph break.
        int replacementCount = doc.Range.Replace(multipleLineBreaks, "\n");

        // Validate that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("No line break sequences were replaced.");

        // Save the modified document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Optionally, write a simple confirmation to the console.
        Console.WriteLine($"Replacements made: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
