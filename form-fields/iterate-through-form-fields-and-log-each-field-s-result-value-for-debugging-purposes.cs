using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a DocumentBuilder for inserting form fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field.
            // Parameters: name, type, format, default text, max length (0 = unlimited).
            builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default Text", 0);
            builder.Writeln(); // Add a line break.

            // Insert a checkbox form field.
            // Parameters: name, isChecked, size (0 = default size).
            builder.InsertCheckBox("CheckBox1", false, 0);
            builder.Writeln();

            // Insert a dropdown (combo box) form field.
            // Parameters: name, list of items, selected index.
            string[] items = { "Option1", "Option2", "Option3" };
            builder.InsertComboBox("DropDown1", items, 0);
            builder.Writeln();

            // Save the document with the created form fields.
            const string outputPath = "FormFieldsExample.docx";
            doc.Save(outputPath);

            // Load the document (optional here, we already have it in memory).
            Document loadedDoc = new Document(outputPath);

            // Ensure that at least one form field exists.
            if (loadedDoc.Range.FormFields.Count == 0)
                throw new InvalidOperationException("The document does not contain any form fields.");

            // Iterate through all form fields and log their Result values.
            foreach (FormField field in loadedDoc.Range.FormFields)
            {
                // Guard against null field (should not happen, but added for safety).
                if (field == null)
                    continue;

                // Log field name and its Result value.
                Console.WriteLine($"Field Name: {field.Name}, Result: {field.Result}");
            }

            // Save the document after reading (no modifications made, but rule requires saving).
            const string finalPath = "FormFieldsExample_Processed.docx";
            loadedDoc.Save(finalPath);
        }
    }
}
