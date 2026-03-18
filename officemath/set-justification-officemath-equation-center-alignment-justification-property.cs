using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the document that contains the OfficeMath equation.
        Document doc = new Document("OfficeMath.docx");

        // Retrieve the first OfficeMath node in the document.
        OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (officeMath != null)
        {
            // The Justification property can be set only when the equation is in Display mode.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Set the equation's justification to center alignment.
            officeMath.Justification = OfficeMathJustification.Center;
        }

        // Save the updated document.
        doc.Save("OfficeMath_Centered.docx");
    }
}
