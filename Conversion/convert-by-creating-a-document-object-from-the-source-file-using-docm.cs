using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourceFile = @"C:\Docs\SourceDocument.docm";

        // Load the DOCM document. The LoadOptions explicitly specify the DOCM format.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Docm, "", "");

        // Create a Document object from the source file using the specified load options.
        Document doc = new Document(sourceFile, loadOptions);

        // The document is now loaded and ready for further processing.
    }
}
