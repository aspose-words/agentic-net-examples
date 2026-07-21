using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with only a Name property.
    public class Person
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
        // Note: No Age property – it is intentionally missing.
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document in memory.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags that reference a missing member (Age).
            builder.Writeln("<<[person.Name]>>"); // Existing member.
            builder.Writeln("<<[person.Age]>>");  // Missing member.

            // 2. Prepare the data source (only Name is present).
            Person data = new Person();

            // 3. Configure the ReportingEngine to allow missing members.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // MissingMemberMessage left as default (empty) so missing members become null literals.

            // 4. Build the report. The root object name must match the tag prefix ("person").
            engine.BuildReport(doc, data, "person");

            // 5. Validate that the missing member was rendered as an empty string.
            // The document should contain only the Name value.
            string resultText = doc.GetText().Trim(); // Remove trailing newlines/spaces.

            // Expected output is just the Name ("John Doe").
            bool isValid = resultText == data.Name;

            // 6. Output validation result.
            Console.WriteLine(isValid
                ? "Success: Missing members rendered as null literals."
                : $"Failure: Unexpected output -> \"{resultText}\"");

            // 7. Save the generated report for inspection.
            doc.Save("Report.docx");
        }
    }
}
