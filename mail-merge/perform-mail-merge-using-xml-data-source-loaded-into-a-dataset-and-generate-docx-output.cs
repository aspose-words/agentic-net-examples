using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new blank document and a builder.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 2. Build the mail‑merge template.
        // -------------------------------------------------
        builder.Writeln("Employee List:");

        // Start of the mail‑merge region.
        builder.InsertField(@"MERGEFIELD TableStart:Employee \* MERGEFORMAT");
        builder.Writeln();

        // Fields that will be repeated for each employee.
        builder.InsertField(@"MERGEFIELD Name \* MERGEFORMAT");
        builder.Write("\t"); // simple separator
        builder.InsertField(@"MERGEFIELD Position \* MERGEFORMAT");
        builder.Writeln();

        // End of the mail‑merge region.
        builder.InsertField(@"MERGEFIELD TableEnd:Employee \* MERGEFORMAT");
        builder.Writeln();

        // -------------------------------------------------
        // 3. Load XML data into a DataSet.
        // -------------------------------------------------
        string xmlData = @"
<Root>
    <Employees>
        <Employee>
            <Name>John Doe</Name>
            <Position>Manager</Position>
        </Employee>
        <Employee>
            <Name>Jane Smith</Name>
            <Position>Developer</Position>
        </Employee>
        <Employee>
            <Name>Bob Johnson</Name>
            <Position>Designer</Position>
        </Employee>
    </Employees>
</Root>";

        DataSet dataSet = new DataSet();
        using (StringReader sr = new StringReader(xmlData))
        {
            dataSet.ReadXml(sr);
        }

        // Ensure the DataTable name matches the region name ("Employee").
        if (!dataSet.Tables.Contains("Employee"))
            throw new InvalidOperationException("Expected DataTable 'Employee' not found in the DataSet.");

        // -------------------------------------------------
        // 4. Execute the mail merge with regions.
        // -------------------------------------------------
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.ExecuteWithRegions(dataSet.Tables["Employee"]);

        // -------------------------------------------------
        // 5. Save the merged document.
        // -------------------------------------------------
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // -------------------------------------------------
        // 6. Validate the output file.
        // -------------------------------------------------
        FileInfo fi = new FileInfo(outputPath);
        if (!fi.Exists)
            throw new FileNotFoundException("The merged document was not created.", outputPath);
        if (fi.Length == 0)
            throw new InvalidOperationException("The merged document is empty.");

        // Simple validation: ensure each employee name appears in the file.
        string resultText = doc.GetText();
        if (!resultText.Contains("John Doe") ||
            !resultText.Contains("Jane Smith") ||
            !resultText.Contains("Bob Johnson"))
        {
            throw new InvalidOperationException("Merged content does not contain expected employee data.");
        }

        Console.WriteLine($"Mail merge completed successfully. Output file: {Path.GetFullPath(outputPath)}");
    }
}
