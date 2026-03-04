using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractFormFieldContent
{
    static void Main()
    {
        // Load the WORDML (DOCX) document that contains form fields.
        // Replace the path with the actual location of your document.
        string inputPath = @"C:\Docs\FormFields.docx";
        Document doc = new Document(inputPath);

        // Iterate through all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;
        for (int i = 0; i < formFields.Count; i++)
        {
            FormField field = formFields[i];

            // The Result property contains the text displayed for the form field.
            // For text input fields this is the user‑entered text,
            // for drop‑down fields it is the selected item,
            // and for check boxes it is either "☒" or "☐".
            string fieldContent = field.Result;

            // Output the field name and its content.
            Console.WriteLine($"Field Name: {field.Name}");
            Console.WriteLine($"Field Type: {field.Type}");
            Console.WriteLine($"Field Content: {fieldContent}");
            Console.WriteLine(new string('-', 40));
        }

        // Optionally, save a copy of the document after processing.
        // This demonstrates the required save lifecycle step.
        string outputPath = Path.Combine(Path.GetDirectoryName(inputPath) ?? "", "FormFields_Processed.docx");
        doc.Save(outputPath);
    }
}
