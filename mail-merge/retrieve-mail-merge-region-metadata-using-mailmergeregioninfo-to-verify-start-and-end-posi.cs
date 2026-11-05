using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the region.
        builder.Writeln("Before region:");

        // Insert the start of the mail merge region.
        builder.InsertField("MERGEFIELD TableStart_MyRegion \\* MERGEFORMAT");
        builder.Writeln();

        // Insert merge fields that belong to the region.
        builder.InsertField("MERGEFIELD Name \\* MERGEFORMAT");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Value \\* MERGEFORMAT");
        builder.Writeln();

        // Insert the end of the mail merge region.
        builder.InsertField("MERGEFIELD TableEnd_MyRegion \\* MERGEFORMAT");
        builder.Writeln();

        // Add some text after the region.
        builder.Writeln("After region.");

        // Prepare a DataTable that matches the region name and field names.
        DataTable table = new DataTable("MyRegion");
        table.Columns.Add("Name");
        table.Columns.Add("Value");
        table.Rows.Add("First", "1");
        table.Rows.Add("Second", "2");

        // Execute the mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(table);

        // Retrieve region metadata manually (GetRegions is not available in this version).
        var regionInfo = new Dictionary<string, (Field Start, Field End)>(StringComparer.OrdinalIgnoreCase);

        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type != FieldType.FieldMergeField)
                continue;

            string fieldCode = field.GetFieldCode(); // e.g., " MERGEFIELD TableStart_MyRegion \\* MERGEFORMAT"
            if (fieldCode.Contains("TableStart_"))
            {
                string regionName = ExtractRegionName(fieldCode, "TableStart_");
                regionInfo[regionName] = (Start: field, End: null);
            }
            else if (fieldCode.Contains("TableEnd_"))
            {
                string regionName = ExtractRegionName(fieldCode, "TableEnd_");
                if (regionInfo.TryGetValue(regionName, out var tuple))
                {
                    regionInfo[regionName] = (Start: tuple.Start, End: field);
                }
                else
                {
                    // End field found before start – store with null start.
                    regionInfo[regionName] = (Start: null, End: field);
                }
            }
        }

        // Report region metadata.
        foreach (var kvp in regionInfo)
        {
            string name = kvp.Key;
            Field startField = kvp.Value.Start;
            Field endField = kvp.Value.End;

            Console.WriteLine($"Region name: {name}");
            Console.WriteLine($"Start field code: {(startField != null ? startField.GetFieldCode() : "Not found")}");
            Console.WriteLine($"End field code: {(endField != null ? endField.GetFieldCode() : "Not found")}");
        }

        // Validation: ensure merged values are present.
        string docText = doc.GetText();
        if (!docText.Contains("First") || !docText.Contains("Second"))
        {
            throw new InvalidOperationException("Merged data was not found in the document.");
        }

        // Save the merged document.
        const string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Merged document saved to '{outputPath}'.");
    }

    private static string ExtractRegionName(string fieldCode, string marker)
    {
        // fieldCode example: " MERGEFIELD TableStart_MyRegion \\* MERGEFORMAT"
        int startIdx = fieldCode.IndexOf(marker, StringComparison.Ordinal);
        if (startIdx < 0)
            return string.Empty;

        startIdx += marker.Length;
        int endIdx = fieldCode.IndexOf(' ', startIdx);
        if (endIdx < 0)
            endIdx = fieldCode.Length;

        return fieldCode.Substring(startIdx, endIdx - startIdx);
    }
}
