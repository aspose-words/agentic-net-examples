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
        builder.Writeln("This  is  a  sample  document  with  double  spaces.");
        builder.Writeln("Another  line  with  double  spaces  here.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches two or more consecutive spaces.
        Regex doubleSpaceRegex = new Regex(@" {2,}");

        // Perform the replacement: double spaces -> single space.
        int replacedCount = loaded.Range.Replace(doubleSpaceRegex, " ", new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No double spaces were found to replace.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
