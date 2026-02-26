using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Text placeholder in the document that will be replaced.
        string placeholder = "[INSERT_HERE]";

        // Content to insert at the placeholder location.
        string newContent = "This is the inserted content.";

        // Perform the replacement within the document's range.
        // The method returns the number of replacements made (optional to use).
        int replacementsMade = doc.Range.Replace(placeholder, newContent);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
