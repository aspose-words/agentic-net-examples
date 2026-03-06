using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document that will serve as a DOT template.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name (empty), defaultValue (unchecked), checkedValue (unchecked), size (0 = auto).
        builder.InsertCheckBox(string.Empty, false, false, 0);

        // Save the document as a Word template (.dot).
        doc.Save("CheckboxTemplate.dot", SaveFormat.Dot);
    }
}
