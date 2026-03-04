using System;
using Aspose.Words;

class AppendDocumentExample
{
    static void Main()
    {
        // Path to the source document that will be appended.
        string sourcePath = "Source.docx";

        // Path to the destination document. This can be an existing file or a new blank document.
        string destinationPath = "Destination.docx";

        // Load the source document from the file system.
        Document srcDoc = new Document(sourcePath);

        // Load the destination document. If the file does not exist, a new blank document will be created.
        Document dstDoc = new Document(destinationPath);

        // Append the source document to the end of the destination document.
        // KeepSourceFormatting preserves the original formatting of the source document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to a new file.
        string outputPath = "CombinedResult.docx";
        dstDoc.Save(outputPath);
    }
}
