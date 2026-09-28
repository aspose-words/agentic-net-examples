using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the start of a mail merge region named "Employees".
        builder.InsertField("TableStart:Employees", null);
        builder.Writeln();

        // Insert a field that will be merged inside the region.
        builder.InsertField("Name", null);
        builder.Writeln();

        // Insert the end of the mail merge region.
        builder.InsertField("TableEnd:Employees", null);
        builder.Writeln();

        // -----------------------------------------------------------------
        // Retrieve metadata about mail merge regions manually.
        // -----------------------------------------------------------------
        // The older Aspose.Words versions may not expose MailMerge.GetRegions().
        // Therefore we scan the document fields for TableStart/TableEnd markers
        // and calculate start/end positions and field counts ourselves.
        // -----------------------------------------------------------------

        // Dictionaries to hold start and end field indexes for each region name.
        var startIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var endIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all fields in the document.
        int fieldIndex = 0;
        foreach (Field field in doc.Range.Fields)
        {
            // Get the raw field code (e.g., "TableStart:Employees").
            string code = field.GetFieldCode();

            if (code.StartsWith("TableStart:", StringComparison.OrdinalIgnoreCase))
            {
                string regionName = code.Substring("TableStart:".Length).Trim();
                startIndexes[regionName] = fieldIndex;
            }
            else if (code.StartsWith("TableEnd:", StringComparison.OrdinalIgnoreCase))
            {
                string regionName = code.Substring("TableEnd:".Length).Trim();
                endIndexes[regionName] = fieldIndex;
            }

            fieldIndex++;
        }

        // Output region information.
        foreach (var kvp in startIndexes)
        {
            string name = kvp.Key;
            int start = kvp.Value;
            int end = endIndexes.ContainsKey(name) ? endIndexes[name] : -1;
            int fieldCount = (end > start) ? (end - start - 1) : 0;

            Console.WriteLine($"Region Name: {name}");
            Console.WriteLine($"Start Field Index: {start}");
            Console.WriteLine($"End Field Index: {end}");
            Console.WriteLine($"Field Count: {fieldCount}");
            Console.WriteLine();
        }
    }
}
