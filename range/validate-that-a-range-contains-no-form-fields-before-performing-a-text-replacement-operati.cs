using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text that includes a placeholder to be replaced later.
        builder.Writeln("Hello _Name_! This is a sample document.");

        // Uncomment the line below to add a form field and see the validation skip the replacement.
        // builder.InsertField("FORMTEXT", "Enter name");

        // Validate that the document's range does not contain any form fields.
        bool containsFormFields = doc.Range.FormFields.Count > 0;

        if (!containsFormFields)
        {
            // No form fields found – perform the text replacement.
            int replaced = doc.Range.Replace("_Name_", "John Doe");
            Console.WriteLine($"Replacements performed: {replaced}");
        }
        else
        {
            // Form fields are present – skip the replacement.
            Console.WriteLine("The range contains form fields; replacement was not performed.");
        }

        // Save the resulting document to the local file system.
        doc.Save("Output.docx");
    }
}
