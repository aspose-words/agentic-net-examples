using System;
using Aspose.Words;
using Aspose.Words.Fields;

class RetrieveFormFieldResult
{
    static void Main()
    {
        // Load an existing Word document that contains form fields.
        // Replace the path with the actual location of your document.
        Document doc = new Document(@"C:\Docs\SampleForm.docx");

        // Specify the name of the form field you want to retrieve.
        // The name is case‑insensitive.
        string fieldName = "MyTextInput";

        // Access the form field by its bookmark name using the FormFields collection indexer.
        FormField formField = doc.Range.FormFields[fieldName];

        // Check that the field exists (null is returned if not found).
        if (formField != null)
        {
            // Read the Result property, which holds the current value of the form field.
            string fieldResult = formField.Result;

            // Output the result to the console.
            Console.WriteLine($"Form field \"{fieldName}\" result: {fieldResult}");
        }
        else
        {
            Console.WriteLine($"Form field \"{fieldName}\" was not found in the document.");
        }

        // Optionally, save the document after any modifications.
        // doc.Save(@"C:\Docs\SampleForm_Updated.docx");
    }
}
