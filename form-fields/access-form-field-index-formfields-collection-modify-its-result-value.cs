using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing document that contains form fields.
            // Replace "Input.docx" with the path to your source document.
            Document doc = new Document("Input.docx");

            // Get the collection of form fields in the document.
            FormFieldCollection formFields = doc.Range.FormFields;

            // Ensure the collection has at least one form field.
            if (formFields.Count == 0)
            {
                Console.WriteLine("The document does not contain any form fields.");
                return;
            }

            // Access a form field by its zero‑based index.
            // For example, modify the first form field (index 0).
            int fieldIndex = 0; // Change this index as needed.
            FormField field = formFields[fieldIndex];

            if (field == null)
            {
                Console.WriteLine($"No form field found at index {fieldIndex}.");
                return;
            }

            // Set the Result property to a new value.
            // For a text input field this changes the displayed text.
            // For a checkbox field use "1" (checked) or "0" (unchecked).
            // For a dropdown field set the text of the desired item.
            field.Result = "New value";

            // Save the modified document.
            // Replace "Output.docx" with the desired output path.
            doc.Save("Output.docx");
        }
    }
}
