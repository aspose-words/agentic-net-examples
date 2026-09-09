using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package reference
using Newtonsoft.Json; // Required package reference

public class Program
{
    public static void Main()
    {
        // Paths for the temporary input and output files.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document containing double spaces.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is  a  sample  text  with  double  spaces.");
        builder.Writeln("Another  line  with  double  spaces.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document we just created.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // Replace any occurrence of two or more spaces with a single space.
        // -----------------------------------------------------------------
        Regex doubleSpacePattern = new Regex(@" {2,}");
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(doubleSpacePattern, " ", options);

        // Validate that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);
    }
}
