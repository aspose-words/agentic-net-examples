using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.BuildingBlocks;

class ReplaceOfficeMathWithPlaceholder
{
    static void Main()
    {
        // Load the DOCM document.
        // The Document constructor is the prescribed way to load a file.
        Document doc = new Document("Input.docm");

        // Collect all OfficeMath nodes in the document.
        // GetChildNodes returns a live collection; we copy it to a list to avoid modification issues while iterating.
        List<OfficeMath> equations = new List<OfficeMath>();
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
            equations.Add(om);

        // Replace each OfficeMath node with a placeholder text.
        foreach (OfficeMath om in equations)
        {
            // Insert a Run containing the placeholder before the OfficeMath node.
            // The Run is created with the same document as its owner.
            om.ParentNode.InsertBefore(new Run(doc, "[Equation]"), om);

            // Remove the original OfficeMath node from the document.
            om.Remove();
        }

        // Save the modified document.
        // The Document.Save method is the prescribed way to persist the file.
        doc.Save("Output.docx");
    }
}
