using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve the form field by its name (replace "MyTextField" with the actual field name).
        FormField formField = doc.Range.FormFields["MyTextField"];
        if (formField != null)
        {
            // Set the value of the form field.
            formField.Result = "New value";
        }
        else
        {
            Console.WriteLine("Form field not found.");
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
