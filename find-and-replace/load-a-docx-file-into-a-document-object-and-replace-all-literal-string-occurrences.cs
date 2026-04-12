using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FindAndReplaceExample
{
    public static void Main()
    {
        // Define file names in the program's working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample_input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample_output.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX file with known content.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Hello Aspose.Words! Hello world!");
        sampleDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the created document.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Step 3: Perform a literal string find‑and‑replace.
        // -----------------------------------------------------------------
        string findText = "Hello";
        string replaceText = "Hi";

        int replacementsMade = doc.Range.Replace(findText, replaceText);
        if (replacementsMade == 0)
        {
            throw new InvalidOperationException($"No occurrences of \"{findText}\" were found to replace.");
        }

        // -----------------------------------------------------------------
        // Step 4: Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // Optional: Inform the user that the operation succeeded.
        Console.WriteLine($"Replaced {replacementsMade} occurrence(s) of \"{findText}\" with \"{replaceText}\".");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
