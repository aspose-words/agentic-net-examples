using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("input.docx");

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save("output.docx");
    }
}
