using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Iterate through all form fields in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Identify combo box (drop‑down) form fields.
            if (field.Type == FieldType.FieldFormDropDown)
            {
                // Set the selected index (zero‑based) to the desired item.
                // For example, select the third item in the list.
                field.DropDownSelectedIndex = 2;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
