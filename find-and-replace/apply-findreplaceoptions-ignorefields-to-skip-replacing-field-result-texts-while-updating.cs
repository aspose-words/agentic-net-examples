using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.IO;

public class FindReplaceIgnoreFieldsExample
{
    public static void Main()
    {
        // Create a sample document with a MERGEFIELD that contains the text "OldName".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document before replacement:");
        // Insert a MERGEFIELD. The field code contains "OldName".
        builder.InsertField("MERGEFIELD OldName \\* MERGEFORMAT");
        builder.Writeln(); // Add a line break.
        // Add normal text that also contains "OldName" to demonstrate that it will be replaced.
        builder.Writeln("This is a normal paragraph with OldName inside.");

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Configure FindReplaceOptions to ignore field result texts.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Replace "OldName" with "NewName". This will affect the field code but not the displayed result.
        int replacedCount = loadedDoc.Range.Replace("OldName", "NewName", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none occurred.");

        // Update fields so that the field result reflects the new field code.
        loadedDoc.UpdateFields();

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Simple validation: ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
