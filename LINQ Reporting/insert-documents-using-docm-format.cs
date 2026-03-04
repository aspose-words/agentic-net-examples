using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document that will be saved as DOCM.
        Document destination = new Document();

        // Use DocumentBuilder to add content and to insert another document.
        DocumentBuilder builder = new DocumentBuilder(destination);
        builder.Writeln("Start of the destination document.");

        // Load the source document that we want to insert.
        // Replace the path with the actual location of your source file.
        Document source = new Document("Source.docx");

        // Insert the source document at the current cursor position,
        // preserving the source formatting.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        builder.Writeln("End of the destination document.");

        // Save the combined document as a macro‑enabled DOCM file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        destination.Save("CombinedDocument.docm", saveOptions);
    }
}
