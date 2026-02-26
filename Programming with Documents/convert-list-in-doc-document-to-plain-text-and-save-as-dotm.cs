using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertListToPlainTextAndSaveAsDotm
{
    static void Main()
    {
        // Path to the source DOC file containing the list.
        string sourceDocPath = @"C:\Docs\SourceDocument.doc";

        // Path to the destination DOTM file.
        string destinationDotmPath = @"C:\Docs\ResultDocument.dotm";

        // Load the existing DOC document.
        Document doc = new Document(sourceDocPath);

        // Ensure that list labels are up‑to‑date.
        doc.UpdateListLabels();

        // OPTIONAL: If you need to simplify complex list labels when exporting to plain text,
        // you could save the document to a temporary text file using TxtSaveOptions.
        // Here we demonstrate the option, but the final output is saved as DOTM.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true // Convert complex list symbols to simple ASCII equivalents.
        };
        // Save to a temporary .txt (not required for DOTM output, shown for completeness).
        // string tempTxtPath = @"C:\Docs\TempPlainText.txt";
        // doc.Save(tempTxtPath, txtOptions);

        // Save the document as a macro‑enabled template (DOTM).
        doc.Save(destinationDotmPath, SaveFormat.Dotm);
    }
}
