using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Obtain the collection of form fields in the whole document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Remove each form field. Iterate backwards to avoid index shifting.
        for (int i = formFields.Count - 1; i >= 0; i--)
        {
            // Remove the form field at the current index.
            formFields.RemoveAt(i);
        }

        // Save the document after all form fields have been removed.
        doc.Save("Output.docx");
    }
}
