using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder at the end of the document (or any other location as needed).
        builder.MoveToDocumentEnd();

        // Insert a checkbox form field.
        // Parameters: name (empty string), default checked value (false), size (0 = auto).
        builder.InsertCheckBox(string.Empty, false, 0);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
