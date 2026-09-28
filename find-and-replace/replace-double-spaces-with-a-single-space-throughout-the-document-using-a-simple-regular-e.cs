using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing double spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is  a  sample  text  with  double  spaces.");
        builder.Writeln("Another  line  with  double  spaces  here.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regular expression that matches two or more consecutive spaces.
        Regex doubleSpaceRegex = new Regex(@" {2,}");

        // Replace all occurrences of double spaces with a single space.
        int replacedCount = loaded.Range.Replace(doubleSpaceRegex, " ", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one double-space replacement, but none were found.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
