using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

class MailMergeRegionValidator
{
    static void Main()
    {
        const string fileName = "MailMergeRegions.docx";

        // Ensure the sample document exists. If not, create a minimal one with a mail‑merge region.
        if (!System.IO.File.Exists(fileName))
            CreateSampleDocument(fileName);

        // Load the source document that contains mail merge regions.
        Document doc = new Document(fileName);

        // Obtain the full hierarchy of mail merge regions.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // Validate each top‑level region.
        ValidateRegions(doc, hierarchy.Regions, 0);

        // Save the document after validation (optional, e.g., to mark processed regions).
        doc.Save("MailMergeRegions_Validated.docx");
    }

    // Creates a simple document containing a single mail‑merge region for demonstration purposes.
    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Before region text.");

        // Insert a start mail‑merge region field named "Region1".
        builder.InsertField("MERGEFIELD  RegionStart_Region1 \\* MERGEFORMAT");
        builder.Writeln();

        // Insert a couple of merge fields inside the region.
        builder.InsertField("MERGEFIELD  FieldA  \\* MERGEFORMAT");
        builder.Writeln();
        builder.InsertField("MERGEFIELD  FieldB  \\* MERGEFORMAT");
        builder.Writeln();

        // Insert an end mail‑merge region field.
        builder.InsertField("MERGEFIELD  RegionEnd_Region1 \\* MERGEFORMAT");
        builder.Writeln();

        builder.Writeln("After region text.");

        doc.Save(path);
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
        NodeCollection allNodes = doc.GetChildNodes(NodeType.Any, true);
        for (int i = 0; i < allNodes.Count; i++)
        {
            if (allNodes[i] == node)
                return i;
        }
        return -1; // Not found.
    }
}
