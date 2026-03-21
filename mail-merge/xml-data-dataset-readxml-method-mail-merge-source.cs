using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsMailMergeExample
{
    class Program
    {
        static void Main()
        {
            // XML data as a string.
            const string xmlData = @"
<DataSet>
    <Employees>
        <Employee>
            <Name>John</Name>
            <Age>30</Age>
        </Employee>
        <Employee>
            <Name>Jane</Name>
            <Age>25</Age>
        </Employee>
    </Employees>
</DataSet>";

            // Load the XML into a DataSet.
            var dataSet = new DataSet();
            using (var reader = new StringReader(xmlData))
            {
                dataSet.ReadXml(reader);
            }

            // Create a Word document template with a mail‑merge region.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Employee List:");
            builder.StartTable();

            // Header row (optional, not part of the mail‑merge region).
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("Age");
            builder.EndRow();

            // Data row with merge fields. The table name (Employees) is inferred from the field prefixes.
            builder.InsertCell();
            builder.InsertField("MERGEFIELD Employees.Name", null);
            builder.InsertCell();
            builder.InsertField("MERGEFIELD Employees.Age", null);
            builder.EndRow();

            builder.EndTable();

            // Perform mail merge using the DataSet.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the merged document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedOutput.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Merged document saved to: {outputPath}");
        }
    }
}
