using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOTM document.
        Document doc = new Document("Template.dotm");

        // Apply protection that allows only form fields to be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document as a DOTM file, preserving the form fields.
        doc.Save("ProtectedTemplate.dotm");
    }
}
