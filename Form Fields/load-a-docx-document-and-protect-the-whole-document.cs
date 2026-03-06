using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("input.docx");

        // Protect the whole document (read‑only, no password).
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        doc.Save("output.docx");
    }
}
