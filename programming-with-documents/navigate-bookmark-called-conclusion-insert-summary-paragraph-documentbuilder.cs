using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing document that contains the "Conclusion" bookmark.
        Document doc = new Document("Input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the bookmark named "Conclusion".
        // The method returns true if the bookmark exists; otherwise false.
        bool bookmarkFound = builder.MoveToBookmark("Conclusion");
        if (!bookmarkFound)
        {
            // If the bookmark is missing, exit or handle the situation as needed.
            return;
        }

        // Insert a new paragraph with the summary text at the bookmark location.
        builder.Writeln("This is the summary paragraph that concludes the document.");

        // Save the updated document to a new file.
        doc.Save("Output.docx");
    }
}
