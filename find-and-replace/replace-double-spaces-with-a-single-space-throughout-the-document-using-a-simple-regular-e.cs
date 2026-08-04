using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with double spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is  a  sample  text  with  double  spaces.");
        builder.Writeln("Another   line   with   triple   spaces.");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regular expression that matches two or more consecutive spaces.
        Regex doubleSpaceRegex = new Regex(@" {2,}");

        // Perform the replacement: replace each match with a single space.
        int replacementCount = loaded.Range.Replace(doubleSpaceRegex, " ", new FindReplaceOptions());

        // Ensure that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one double‑space replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
