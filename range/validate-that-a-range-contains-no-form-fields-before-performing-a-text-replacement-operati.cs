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

        // Add a paragraph that contains the text we intend to replace.
        builder.Writeln("Hello _Name_! This is a sample document.");

        // Insert a checkbox form field to demonstrate the validation logic.
        // The presence of any form field in the target range should prevent replacement.
        builder.InsertCheckBox("Agree", false, 0);

        // Save the original document (optional, for inspection).
        doc.Save("original.docx");

        // Work with the whole‑document range.
        // Use the fully qualified type name to avoid conflict with System.Range.
        Aspose.Words.Range range = doc.Range;

        // Verify that the range does not contain any form fields.
        bool containsFormFields = range.FormFields.Count > 0;

        if (!containsFormFields)
        {
            // No form fields – perform the replacement.
            int replaced = range.Replace("_Name_", "John Doe");
            Console.WriteLine($"Replacements performed: {replaced}");
        }
        else
        {
            // Form fields detected – skip the replacement.
            Console.WriteLine("The range contains form fields; replacement was skipped.");
        }

        // Save the (potentially) modified document.
        doc.Save("output.docx");
    }
}
