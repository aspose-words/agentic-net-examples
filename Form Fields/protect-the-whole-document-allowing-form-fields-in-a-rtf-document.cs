using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectRtfDocument
{
    static void Main()
    {
        // Path to the source RTF document.
        string inputPath = "input.rtf";

        // Path where the protected RTF document will be saved.
        string outputPath = "output.rtf";

        // Load the existing RTF document.
        Document doc = new Document(inputPath);

        // Protect the entire document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Create RTF save options (default options are sufficient here).
        RtfSaveOptions saveOptions = new RtfSaveOptions();

        // Save the protected document as RTF.
        doc.Save(outputPath, saveOptions);
    }
}
