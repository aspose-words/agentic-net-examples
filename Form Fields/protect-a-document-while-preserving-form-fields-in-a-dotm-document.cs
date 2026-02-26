using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectDotm
{
    static void Main()
    {
        // Path to the source DOTM template.
        string inputPath = @"C:\Docs\Template.dotm";

        // Path where the protected DOTM will be saved.
        string outputPath = @"C:\Docs\Template_Protected.dotm";

        // Load the existing DOTM document.
        Document doc = new Document(inputPath);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the document as a DOTM (macro‑enabled template) using DocSaveOptions.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);
        doc.Save(outputPath, saveOptions);
    }
}
