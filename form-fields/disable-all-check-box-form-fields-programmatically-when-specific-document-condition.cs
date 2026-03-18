using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Example condition: the document contains the marker text "DISABLE_CHECKBOXES".
        bool conditionMet = doc.GetText().Contains("DISABLE_CHECKBOXES");

        if (conditionMet)
        {
            // Iterate through all form fields in the document.
            foreach (FormField field in doc.Range.FormFields)
            {
                // Process only checkbox form fields.
                if (field.Type == FieldType.FieldFormCheckBox)
                {
                    // Disable the checkbox so the user cannot change its state.
                    field.Enabled = false;
                }
            }
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
