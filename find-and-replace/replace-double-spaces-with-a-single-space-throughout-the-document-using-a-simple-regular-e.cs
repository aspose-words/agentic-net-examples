using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;      // Required by the task specification
using Newtonsoft.Json;    // Required by the task specification

public class Program
{
    public static void Main()
    {
        // Create a sample document that contains double spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is  a  sample  text  with  double  spaces.");
        builder.Writeln("Another  line  with  double  spaces.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Regular expression that matches two or more consecutive space characters.
        Regex doubleSpaceRegex = new Regex(@" {2,}");

        // Replace all occurrences of the pattern with a single space.
        int replacedCount = loaded.Range.Replace(doubleSpaceRegex, " ", new FindReplaceOptions());

        // Validate that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
