using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourceFile = @"C:\Docs\SourceDocument.docm";

        // Load the DOCM document. The constructor automatically detects the format,
        // but we can explicitly specify LoadOptions with LoadFormat.Docm if desired.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Docm, "", "");
        Document doc = new Document(sourceFile, loadOptions);

        // At this point 'doc' is a fully loaded Document object representing the DOCM file.
        // Further processing can be performed here.
    }
}
