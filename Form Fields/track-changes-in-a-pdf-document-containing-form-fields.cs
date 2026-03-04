using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class TrackChangesInPdfWithFormFields
{
    static void Main()
    {
        // Load a Word document that contains form fields.
        // Replace with the actual path to your source .docx file.
        string inputPath = @"C:\Docs\FormFields.docx";
        Document doc = new Document(inputPath);

        // Begin tracking revisions programmatically.
        // All subsequent changes will be recorded as revision changes.
        doc.StartTrackRevisions("Reviewer");

        // Example modification: change the value of each form field.
        // This will be captured as an insertion revision for the new field content.
        FormFieldCollection formFields = doc.Range.FormFields;
        foreach (FormField field in formFields)
        {
            switch (field.Type)
            {
                case FieldType.FieldFormCheckBox:
                    // Toggle the checkbox state.
                    field.Checked = !field.Checked;
                    break;
                case FieldType.FieldFormDropDown:
                    // Select the next item in the dropdown list, if any.
                    if (field.DropDownItems.Count > 0)
                    {
                        int nextIndex = (field.DropDownSelectedIndex + 1) % field.DropDownItems.Count;
                        field.DropDownSelectedIndex = nextIndex;
                    }
                    break;
                case FieldType.FieldFormTextInput:
                    // Append a suffix to the existing text input.
                    field.Result = field.Result + " (updated)";
                    break;
            }
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Prepare PDF save options to preserve the form fields as interactive PDF fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };

        // Save the document as PDF. Replace with the desired output path.
        string outputPath = @"C:\Docs\FormFields_Tracked.pdf";
        doc.Save(outputPath, pdfOptions);
    }
}
