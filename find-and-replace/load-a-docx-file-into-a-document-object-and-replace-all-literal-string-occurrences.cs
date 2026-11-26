using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FindAndReplaceExample
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Create a sample document with text that will be replaced.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document. Replace the word 'sample' with 'example'.");
        builder.Writeln("Another line containing the word sample.");
        sampleDoc.Save(inputPath);

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Perform a literal string replacement.
        string findText = "sample";
        string replaceText = "example";
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(findText, replaceText, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException($"No occurrences of '{findText}' were found to replace.");
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
