using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOTM (macro‑enabled template) document.
        Document doc = new Document("Template.dotm");

        // Protect the document so that users can only fill in form fields.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document as a DOTM file.
        // Use DocSaveOptions with SaveFormat.Dot to keep the macro‑enabled template format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);
        doc.Save("ProtectedTemplate.dotm", saveOptions);
    }
}
