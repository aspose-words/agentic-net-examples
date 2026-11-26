using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text to the document.
        builder.Writeln("Hello world! This is a sample document.");

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");

        // Validate that the document's range contains no form fields.
        if (doc.Range.FormFields.Count == 0)
        {
            // Perform a simple text replacement because no form fields are present.
            doc.Range.Replace("Hello", "Hi", new FindReplaceOptions());

            // Save the modified document.
            doc.Save(outputPath);
        }
        else
        {
            // If form fields exist, skip replacement and still save the original document.
            doc.Save(outputPath);
        }
    }
}
