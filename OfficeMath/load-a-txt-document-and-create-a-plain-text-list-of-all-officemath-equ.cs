using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source Word document that contains OfficeMath equations.
        const string sourceDocPath = "Input.docx";

        // Path where the whole document will be exported as plain‑text (demonstrates TxtSaveOptions).
        const string exportedTxtPath = "Exported.txt";

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Configure TxtSaveOptions to export OfficeMath objects as plain text.
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions
        {
            OfficeMathExportMode = TxtOfficeMathExportMode.Text
        };

        // Save the entire document to a TXT file using the configured options.
        doc.Save(exportedTxtPath, txtSaveOptions);

        // Load the exported TXT file using PlainTextDocument (demonstrates TxtLoadOptions usage).
        PlainTextDocument plainTxt = new PlainTextDocument(exportedTxtPath);
        Console.WriteLine("Full plain‑text document content:");
        Console.WriteLine(plainTxt.Text);

        // Collect each OfficeMath equation as a plain‑text string.
        List<string> equationTexts = new List<string>();
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Export the individual OfficeMath node to plain text using the same TxtSaveOptions.
            string equation = officeMath.ToString(txtSaveOptions);
            equationTexts.Add(equation);
        }

        // Output the extracted equations.
        Console.WriteLine("\nExtracted OfficeMath equations:");
        foreach (string eq in equationTexts)
            Console.WriteLine(eq);

        // Optionally write the list of equations to a separate text file.
        File.WriteAllLines("Equations.txt", equationTexts);
    }
}
