using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Get the collection that holds all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // ----- Retrieve a form field by index -----
        if (formFields.Count > 0)
        {
            // FormFieldCollection uses zero‑based indexing.
            FormField fieldByIndex = formFields[0];
            Console.WriteLine($"Field at index 0: Name = {fieldByIndex.Name}, Type = {fieldByIndex.Type}");
        }

        // ----- Retrieve a form field by name -----
        string fieldName = "MyCheckBox"; // replace with the actual field name you need
        FormField fieldByName = formFields[fieldName];
        if (fieldByName != null)
        {
            Console.WriteLine($"Field named '{fieldName}': Type = {fieldByName.Type}, Checked = {fieldByName.Checked}");
        }

        // Save the document if any changes were made.
        doc.Save("Output.docx");
    }
}
