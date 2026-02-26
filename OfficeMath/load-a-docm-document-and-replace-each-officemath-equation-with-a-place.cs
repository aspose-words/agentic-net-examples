using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Replacing;

class ReplaceOfficeMathWithPlaceholder
{
    static void Main()
    {
        // Load the DOCM document.
        // If the document contains shapes with EquationXML that should be treated as OfficeMath,
        // you can configure LoadOptions accordingly. Here we use default loading.
        Document doc = new Document("Input.docm");

        // Collect all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate over a copy of the collection because we will modify the document structure.
        List<OfficeMath> officeMathList = new List<OfficeMath>();
        foreach (OfficeMath om in officeMathNodes)
            officeMathList.Add(om);

        // Replace each OfficeMath node with a placeholder Run.
        foreach (OfficeMath om in officeMathList)
        {
            // Create a Run containing the placeholder text.
            Run placeholder = new Run(doc, "[Equation]");

            // Insert the placeholder before the OfficeMath node.
            om.ParentNode.InsertBefore(placeholder, om);

            // Remove the original OfficeMath node.
            om.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
