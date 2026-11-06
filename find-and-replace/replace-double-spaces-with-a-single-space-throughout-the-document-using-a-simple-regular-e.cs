using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Create a new document and add text containing double spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is  a  sample  document.  It  contains  double  spaces  everywhere.");
        builder.Writeln("Another  line  with  double  spaces.");

        // Define a regular expression that matches two or more consecutive spaces.
        Regex doubleSpaceRegex = new Regex(@" {2,}");

        // Perform the replacement: replace each occurrence with a single space.
        int replacedCount = doc.Range.Replace(doubleSpaceRegex, " ", new FindReplaceOptions());

        // Validate that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one double-space replacement, but none were found.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
