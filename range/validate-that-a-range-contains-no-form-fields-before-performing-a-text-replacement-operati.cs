using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a placeholder that we intend to replace.
        builder.Writeln("Dear _Name_, welcome to the company.");

        // --------------------------------------------------------------------
        // Validation: ensure the whole document range does NOT contain any form fields.
        // The FormFields collection is available on the Range object.
        // --------------------------------------------------------------------
        bool hasFormFields = doc.Range.FormFields.Count > 0;

        if (hasFormFields)
        {
            // If form fields are present, we abort the replacement operation.
            Console.WriteLine("The document contains form fields; replacement aborted.");
        }
        else
        {
            // No form fields – safe to perform the replacement.
            int replacements = doc.Range.Replace("_Name_", "John Doe");
            Console.WriteLine($"Replacements made: {replacements}");
        }

        // Save the resulting document to the local file system.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
