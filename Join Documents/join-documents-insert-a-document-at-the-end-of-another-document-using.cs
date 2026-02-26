using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document (the one to which we will append).
        Document destination = new Document("Destination.docx");

        // Load the source document that will be appended.
        Document source = new Document("Source.docx");

        // Append the source document to the end of the destination document,
        // preserving the source document's formatting.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        destination.Save("Combined.docx");
    }
}
