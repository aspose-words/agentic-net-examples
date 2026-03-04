using System;
using Aspose.Words;
using Aspose.Words.Fields;

class FormFieldAccessExample
{
    static void Main()
    {
        // Path to the source DOCX file that contains form fields.
        string inputPath = @"C:\Docs\InputFormFields.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Access the collection of form fields in the whole document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // ----- Retrieve a form field by zero‑based index -----
        // Example: get the first form field (index 0).
        FormField fieldByIndex = formFields[0];
        if (fieldByIndex != null)
        {
            Console.WriteLine("Field by index:");
            Console.WriteLine($"  Name   : {fieldByIndex.Name}");
            Console.WriteLine($"  Type   : {fieldByIndex.Type}");
            Console.WriteLine($"  Result : {fieldByIndex.Result}");
        }
        else
        {
            Console.WriteLine("No form field found at the specified index.");
        }

        // ----- Retrieve a form field by its bookmark/name -----
        // Example: get a field named "MyComboBox".
        string fieldName = "MyComboBox";
        FormField fieldByName = formFields[fieldName];
        if (fieldByName != null)
        {
            Console.WriteLine("\nField by name:");
            Console.WriteLine($"  Name   : {fieldByName.Name}");
            Console.WriteLine($"  Type   : {fieldByName.Type}");
            Console.WriteLine($"  Result : {fieldByName.Result}");
        }
        else
        {
            Console.WriteLine($"\nForm field with name \"{fieldName}\" not found.");
        }

        // Optionally, modify a field (e.g., change the result of a text input field).
        if (fieldByName != null && fieldByName.Type == FieldType.FieldFormTextInput)
        {
            fieldByName.Result = "New value set by code";
        }

        // Save the modified document to a new file.
        string outputPath = @"C:\Docs\OutputFormFields.docx";
        doc.Save(outputPath);
        Console.WriteLine($"\nDocument saved to: {outputPath}");
    }
}
