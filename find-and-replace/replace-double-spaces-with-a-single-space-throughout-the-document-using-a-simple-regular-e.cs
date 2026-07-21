using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing double spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is  a  test  document.  It  contains  double  spaces.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regular expression that matches two or more consecutive spaces.
        Regex doubleSpaceRegex = new Regex(@" {2,}");

        // Replace all occurrences of double (or more) spaces with a single space.
        int replaced = loaded.Range.Replace(doubleSpaceRegex, " ", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
