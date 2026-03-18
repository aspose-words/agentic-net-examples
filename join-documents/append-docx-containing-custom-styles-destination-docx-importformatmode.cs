using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source (with custom styles), destination, and output PDF files.
        string sourcePath = @"C:\Docs\SourceWithCustomStyles.docx";
        string destinationPath = @"C:\Docs\Destination.docx";
        string outputPdfPath = @"C:\Docs\CombinedResult.pdf";

        // Load the documents from disk.
        Document srcDoc = new Document(sourcePath);
        Document dstDoc = new Document(destinationPath);

        // Append the source document to the destination document.
        // UseDestinationStyles ensures that existing styles in the destination are kept.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // Export the merged document to PDF.
        dstDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
