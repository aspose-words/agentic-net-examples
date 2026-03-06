using Aspose.Words;

class Program
{
    static void Main()
    {
        // Paths to the source (to be inserted) and destination documents.
        string dataDir = @"C:\Data\";
        string destPath = dataDir + "Destination.docx";
        string srcPath = dataDir + "Source.docx";

        // Load the destination document (or create a new blank one if you prefer).
        Document destDoc = new Document(); // or new Document(destPath) to load an existing file

        // Create a builder for the destination document and add a bookmark.
        DocumentBuilder builder = new DocumentBuilder(destDoc);
        builder.StartBookmark("InsertHere");
        builder.Write("Content before insertion. ");
        builder.EndBookmark("InsertHere");
        builder.Writeln(" Content after insertion.");

        // Load the source document that will be inserted.
        Document srcDoc = new Document(srcPath);

        // Move the builder's cursor to the bookmark.
        if (builder.MoveToBookmark("InsertHere"))
        {
            // Insert the source document at the bookmark location.
            // KeepSourceFormatting preserves the original formatting of the inserted document.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined document.
        destDoc.Save(destPath);
    }
}
