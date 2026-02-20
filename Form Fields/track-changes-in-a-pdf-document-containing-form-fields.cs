using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the PDF document that contains form fields.
        Document doc = new Document("FormFields.pdf");

        // Turn on change tracking so that all subsequent edits are recorded as revisions.
        doc.TrackRevisions = true;

        // Iterate through all form fields in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // If the field is a text input, change its displayed value.
            if (field.Type == FieldType.FieldFormTextInput)
            {
                // The Result property holds the current contents of the form field.
                field.Result = "New value";
            }
            // If the field is a check box, toggle its checked state.
            else if (field.Type == FieldType.FieldFormCheckBox)
            {
                field.Checked = !field.Checked;
            }
            // Add additional handling for other field types as needed.
        }

        // Save the document back to PDF. The saved file will contain the tracked changes (revisions).
        doc.Save("FormFields_Tracked.pdf");
    }
}
