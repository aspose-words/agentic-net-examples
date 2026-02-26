using System.Reflection;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the MHTML document. Enable conversion of EquationXML shapes to OfficeMath objects.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            ConvertShapeToOfficeMath = true
        };
        Document doc = new Document("input.mht", loadOptions);

        // Iterate through all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in mathNodes)
        {
            // MathObjectType is read‑only; use reflection to set it to Matrix.
            // The internal field name may differ between versions; adjust if necessary.
            FieldInfo field = typeof(OfficeMath).GetField("_mathObjectType", BindingFlags.Instance | BindingFlags.NonPublic);
            if (field != null)
            {
                // Assume MathObjectType.Matrix exists in the enum.
                field.SetValue(officeMath, MathObjectType.Matrix);
            }
        }

        // Save the modified document back to MHTML.
        doc.Save("output.mht");
    }
}
