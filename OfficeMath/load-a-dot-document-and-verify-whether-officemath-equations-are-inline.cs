using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class OfficeMathInlineChecker
{
    // Checks whether all top‑level OfficeMath objects in a DOT document are displayed inline.
    // Returns true if every OfficeMath node has DisplayType == Inline; otherwise false.
    public static bool AreAllOfficeMathInline(string dotFilePath)
    {
        // Load the DOT (Word template) document.
        // No special LoadOptions are required for this check.
        Document doc = new Document(dotFilePath);

        // Retrieve all OfficeMath nodes in the document (including nested ones).
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // If there are no OfficeMath nodes, we consider the document as having no equations to verify.
        if (mathNodes.Count == 0)
            return true;

        // Iterate through each OfficeMath node and verify its DisplayType.
        foreach (OfficeMath officeMath in mathNodes)
        {
            // For nested OfficeMath objects the DisplayType is always Inline by definition,
            // but we still check the property for completeness.
            if (officeMath.DisplayType != OfficeMathDisplayType.Inline)
                return false; // Found a non‑inline equation.
        }

        // All equations are inline.
        return true;
    }

    // Example usage.
    static void Main()
    {
        string dotPath = @"C:\Docs\Template.dot"; // Path to the DOT file.

        bool allInline = AreAllOfficeMathInline(dotPath);

        Console.WriteLine(allInline
            ? "All OfficeMath equations are inline."
            : "There are OfficeMath equations that are not inline.");
    }
}
