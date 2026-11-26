using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document. Hello World! This phrase will be replaced.");

        // Phrase to search and its replacement.
        string findText = "Hello World";
        string replaceText = "Hi Universe";

        // Replace the phrase within the whole-document range.
        doc.Range.Replace(findText, replaceText, new FindReplaceOptions());

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the replacement occurred.
        bool succeeded = doc.Range.Text.Contains(replaceText);
        Console.WriteLine(succeeded ? "Replacement succeeded." : "Replacement failed.");
    }
}
