using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the MHTML document. If the source contains EquationXML shapes and you want them
        // converted to OfficeMath objects, enable the corresponding option.
        LoadOptions loadOptions = new LoadOptions
        {
            ConvertShapeToOfficeMath = true
        };
        Document doc = new Document("input.mht", loadOptions);

        // Iterate through all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // The MathObjectType property is read‑only; Aspose.Words does not provide an API
            // to change the type of an existing OfficeMath object (e.g., from OMath to Matrix).
            // Therefore we cannot directly set it to MathObjectType.Matrix.
            // If a transformation is required, the node would need to be replaced with a new
            // OfficeMath object of the desired type, which is not supported by the current API.
            Console.WriteLine($"Current MathObjectType: {officeMath.MathObjectType}");
        }

        // Save the (unchanged) document back to MHTML.
        doc.Save("output.mht", SaveFormat.Mhtml);
    }
}
