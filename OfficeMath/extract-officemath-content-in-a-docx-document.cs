using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Saving;

class ExtractOfficeMath
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve all OfficeMath nodes in the document.
        var officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true)
                                 .Cast<OfficeMath>()
                                 .ToList();

        // Extract plain‑text representation of each OfficeMath node.
        using (StringWriter writer = new StringWriter())
        {
            foreach (var om in officeMathNodes)
            {
                string text = om.GetText().Trim();
                writer.WriteLine(text);
            }

            // Save the extracted OfficeMath text to a separate file.
            File.WriteAllText("OfficeMathExtracted.txt", writer.ToString());
        }

        // Additionally, save the whole document with OfficeMath exported as LaTeX.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.OfficeMathExportMode = TxtOfficeMathExportMode.Latex;
        doc.Save("OfficeMathLatex.txt", txtOptions);
    }
}
