using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load an existing Word document that contains an OfficeMath equation.
        // Replace the path with the actual location of your source document.
        Document doc = new Document("OfficeMath.docx");

        // Retrieve the first OfficeMath node in the document.
        // The GetChild method searches the document tree for a node of the specified type.
        OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

        // Change the display format so the equation appears on its own line.
        officeMath.DisplayType = OfficeMathDisplayType.Display;

        // Optionally set the justification of the displayed equation.
        // This must be done after setting DisplayType.
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the modified document to a new file.
        // Replace the path with the desired output location.
        doc.Save("OfficeMath_Display.docx");
    }
}
