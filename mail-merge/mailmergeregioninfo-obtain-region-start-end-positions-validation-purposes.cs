using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

class MailMergeRegionValidator
{
    static void Main()
    {
        // Load the source document that contains mail merge regions.
        Document doc = new Document("MailMergeRegions.docx");

        // Obtain the full hierarchy of mail merge regions.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // Validate each top‑level region.
        ValidateRegions(doc, hierarchy.Regions, 0);
        
        // Save the document after validation (optional, e.g., to mark processed regions).
        doc.Save("MailMergeRegions_Validated.docx");
    }

    // Recursively validates a collection of regions.
    private static void ValidateRegions(Document doc, IList<MailMergeRegionInfo> regions, int indentLevel)
    {
        string indent = new string(' ', indentLevel * 4);
        foreach (MailMergeRegionInfo region in regions)
        {
            // Region name.
            Console.WriteLine($"{indent}Region: {region.Name}");

            // Start and end merge fields that delimit the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField   = region.EndField;

            // Output the field names for verification.
            Console.WriteLine($"{indent}    Start field name: {startField?.FieldName}");
            Console.WriteLine($"{indent}    End field   name: {endField?.FieldName}");

            // Obtain the positions (node indices) of the start and end fields within the document.
            // This can be useful for custom validation logic.
            int startPos = GetNodeIndex(doc, startField?.Start);
            int endPos   = GetNodeIndex(doc, endField?.Start);
            Console.WriteLine($"{indent}    Start field position (node index): {startPos}");
            Console.WriteLine($"{indent}    End field   position (node index): {endPos}");

            // Recursively validate any nested regions.
            if (region.Regions != null && region.Regions.Count > 0)
                ValidateRegions(doc, region.Regions, indentLevel + 1);
        }
    }

    // Returns the zero‑based index of a node within the document's node collection.
    private static int GetNodeIndex(Document doc, Node node)
    {
        if (node == null) return -1;
        // The document's GetChildNodes method returns a collection of all nodes of a given type.
        // Here we use NodeType.Any to get a flat list of all nodes in document order.
        NodeCollection allNodes = doc.GetChildNodes(NodeType.Any, true);
        for (int i = 0; i < allNodes.Count; i++)
        {
            if (allNodes[i] == node)
                return i;
        }
        return -1; // Not found.
    }
}
