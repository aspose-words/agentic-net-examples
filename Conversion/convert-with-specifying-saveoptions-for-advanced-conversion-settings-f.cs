using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDocmConversion
{
    class Program
    {
        static void Main()
        {
            // Load the source document (any supported format, e.g., DOCX).
            Document doc = new Document("input.docx");

            // Create save options for DOCM format with advanced settings.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm)
            {
                // Example: use maximum compression for the OOXML package.
                CompressionLevel = CompressionLevel.Maximum,

                // Example: set a password to protect the output DOCM file.
                Password = "myPassword",

                // Example: update fields before saving.
                UpdateFields = true,

                // Example: embed the Aspose.Words generator name.
                ExportGeneratorName = true,

                // Example: enable pretty formatting of the XML parts.
                PrettyFormat = true
            };

            // Save the document as a macro‑enabled DOCM file using the specified options.
            doc.Save("output.docm", saveOptions);
        }
    }
}
