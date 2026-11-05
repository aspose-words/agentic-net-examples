using System;
using Aspose.Words;
using Aspose.Words.Replacing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Words.Range;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This is a sample document.");

        // Save the source document locally.
        const string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Load the document to work with its range.
        Document loadedDoc = new Document(sourcePath);
        AsposeRange range = loadedDoc.Range;

        // Validate that the range contains no form fields.
        if (range.FormFields.Count == 0)
        {
            // Since there are no form fields, perform a text replacement.
            range.Replace("Hello", "Hi", new FindReplaceOptions(FindReplaceDirection.Forward));
        }

        // Save the resulting document.
        const string outputPath = "Result.docx";
        loadedDoc.Save(outputPath);
    }
}
