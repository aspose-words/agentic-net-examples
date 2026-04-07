using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------- Data source --------------------
        DataTable people = new();
        people.TableName = "People";
        people.Columns.Add("Name");
        people.Columns.Add("Age");
        people.Rows.Add(new object[] { "Alice", 30 });
        people.Rows.Add(new object[] { "Bob", 45 });
        people.Rows.Add(new object[] { "Charlie", 28 });

        // -------------------- Build template --------------------
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in People]>>");

        // Start a table that will be repeated for each person.
        Table table = builder.StartTable();

        // ----- Name row -----
        builder.InsertCell();                     // Header cell
        builder.Write("Name:");
        builder.InsertCell();                     // Value cell

        // Insert a plain‑text content control (SDT) and write the LINQ tag inside it.
        StructuredDocumentTag nameSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        builder.Writeln("<<[person.Name]>>");
        builder.EndRow();

        // ----- Age row -----
        builder.InsertCell();                     // Header cell
        builder.Write("Age:");
        builder.InsertCell();                     // Value cell

        StructuredDocumentTag ageSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        builder.Writeln("<<[person.Age]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // -------------------- Generate report --------------------
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None; // default options
        bool success = engine.BuildReport(template, people, "People");

        // Save the generated report.
        template.Save("Report.docx");
    }
}
