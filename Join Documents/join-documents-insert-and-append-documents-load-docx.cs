using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Define file locations.
        string folder = @"C:\Docs\";
        string firstDocPath = Path.Combine(folder, "First.docx");
        string secondDocPath = Path.Combine(folder, "Second.docx");
        string outputPath = Path.Combine(folder, "Combined.docx");

        // Load the source documents.
        Document firstDoc = new Document(firstDocPath);
        Document secondDoc = new Document(secondDocPath);

        // Create a new blank document that will hold the result.
        Document resultDoc = new Document();

        // Use DocumentBuilder to insert the first document at the end of the result.
        DocumentBuilder builder = new DocumentBuilder(resultDoc);
        builder.MoveToDocumentEnd();
        builder.InsertDocument(firstDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert a page break to separate the documents.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the second document after the page break.
        builder.InsertDocument(secondDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        resultDoc.Save(outputPath);
    }
}
