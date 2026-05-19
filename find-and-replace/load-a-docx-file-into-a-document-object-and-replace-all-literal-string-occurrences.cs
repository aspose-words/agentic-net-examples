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
        // Create a sample DOCX file with known content.
        string inputPath = "input.docx";
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document. Replace the word TARGET wherever it appears. TARGET appears twice.");
        sampleDoc.Save(inputPath);

        // Load the created document.
        Document loadedDoc = new Document(inputPath);

        // Perform a literal string find-and-replace.
        string findText = "TARGET";
        string replaceText = "REPLACED";
        int replacementCount = loadedDoc.Range.Replace(findText, replaceText, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException($"Expected at least one occurrence of \"{findText}\" to be replaced.");

        // Save the modified document.
        string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
