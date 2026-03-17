using System;
using Aspose.Words;
using Aspose.Words.Loading;

class TableFormattingExample
{
    // Entry point required by the C# compiler.
    static void Main(string[] args)
    {
        LoadDocumentPreservingTableFormatting();
    }

    public static void LoadDocumentPreservingTableFormatting()
    {
        // No special LoadOptions are required – Aspose.Words preserves the original
        // table formatting by default when a document is loaded.
        // If you need to customise loading (e.g., password, encoding), create a LoadOptions
        // instance and set the appropriate properties, but do NOT use a non‑existent
        // PreserveFormatting property.
        LoadOptions loadOptions = new LoadOptions();

        // Load the document using the (default) options.
        Document doc = new Document("Input.docx", loadOptions);

        // Save the document (optional) to verify that the table formatting remains unchanged.
        doc.Save("Output.docx");
    }
}
