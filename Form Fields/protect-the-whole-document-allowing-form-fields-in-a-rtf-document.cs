using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class ProtectRtfWithFormFields
{
    static void Main()
    {
        // Path to the source RTF document.
        string inputPath = @"C:\Docs\SourceDocument.rtf";

        // Load the RTF document using default load options.
        Document doc = new Document(inputPath, new RtfLoadOptions());

        // Apply protection that allows only form field editing.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Prepare save options for RTF format (optional, can be omitted for defaults).
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Ensure the document is saved as RTF.
            SaveFormat = SaveFormat.Rtf
        };

        // Path to the protected output RTF document.
        string outputPath = @"C:\Docs\ProtectedDocument.rtf";

        // Save the protected document.
        doc.Save(outputPath, saveOptions);
    }
}
