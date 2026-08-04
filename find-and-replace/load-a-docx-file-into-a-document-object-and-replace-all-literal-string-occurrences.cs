using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required by the category rules
using Newtonsoft.Json;        // Required by the category rules

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file with text that contains the target string.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";
        const string searchText = "old";
        const string replaceText = "new";

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"This is an {searchText} value.");
        builder.Writeln($"Another {searchText} appears here.");
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Perform a literal string replace.
        int replacedCount = loaded.Range.Replace(searchText, replaceText, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
