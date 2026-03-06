using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the TXT file into a Word document.
        Document doc = new Document("input.txt");

        // Protect the entire document, allowing only form field editing.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save("output.docx");
    }
}
