using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder for template creation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple template with a mail‑merge region named "Employees".
        builder.Writeln("Report Header");
        builder.InsertField("TableStart:Employees", null);
        builder.Writeln("\tName: ");
        builder.InsertField("MERGEFIELD Name", null);
        builder.Writeln("\tAge: ");
        builder.InsertField("MERGEFIELD Age", null);
        builder.InsertField("TableEnd:Employees", null);
        builder.Writeln("Report Footer");

        // -----------------------------------------------------------------
        // Retrieve and display region metadata (start and end field codes).
        // Aspose.Words older versions may not expose GetRegions(), so we locate
        // the TableStart/TableEnd fields manually.
        // -----------------------------------------------------------------
        FieldCollection fields = doc.Range.Fields;
        for (int i = 0; i < fields.Count; i++)
        {
            Field startField = fields[i];
            string startCode = startField.GetFieldCode();

            if (startCode.StartsWith("TableStart:", StringComparison.OrdinalIgnoreCase))
            {
                // Extract region name.
                string regionName = startCode.Substring("TableStart:".Length).Trim();

                // Find the matching TableEnd field.
                Field endField = null;
                for (int j = i + 1; j < fields.Count; j++)
                {
                    Field possibleEnd = fields[j];
                    string endCode = possibleEnd.GetFieldCode();
                    if (endCode.StartsWith("TableEnd:", StringComparison.OrdinalIgnoreCase) &&
                        endCode.Substring("TableEnd:".Length).Trim().Equals(regionName, StringComparison.OrdinalIgnoreCase))
                    {
                        endField = possibleEnd;
                        break;
                    }
                }

                Console.WriteLine($"Region Name: {regionName}");
                Console.WriteLine($"  Start Field Code: {startCode}");
                Console.WriteLine($"  End Field Code: {(endField != null ? endField.GetFieldCode() : "Not found")}");
            }
        }

        // Prepare data for the region merge.
        DataTable employeeTable = new DataTable("Employees");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Age", typeof(int));

        employeeTable.Rows.Add("Alice", 30);
        employeeTable.Rows.Add("Bob", 45);
        employeeTable.Rows.Add("Charlie", 28);

        // Execute the region‑based mail merge.
        doc.MailMerge.ExecuteWithRegions(employeeTable);

        // Save the merged document.
        const string outputPath = "Merged.docx";
        doc.Save(outputPath);

        // Validate that the merged document contains content.
        string mergedText = doc.GetText();
        if (string.IsNullOrWhiteSpace(mergedText))
        {
            throw new InvalidOperationException("Merged document is empty despite providing data.");
        }

        Console.WriteLine($"Mail merge completed successfully. Output saved to '{outputPath}'.");
    }
}
