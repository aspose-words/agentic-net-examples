using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a template document that contains a reference to a
        //    member that will be missing in the data source.
        // ------------------------------------------------------------
        Document template = new Document();                     // create a new blank document
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<[person.Name]>>");                  // template expression referencing a missing member

        // Save the template as a DOCM (macro‑enabled) file.
        const string templatePath = "Template.docm";
        template.Save(templatePath, SaveFormat.Docm);          // load/save rule applied

        // ------------------------------------------------------------
        // 2. Load the DOCM template back into a Document object.
        // ------------------------------------------------------------
        Document doc = new Document(templatePath);             // load rule applied

        // ------------------------------------------------------------
        // 3. Prepare a data source that does NOT contain the 'Name' property.
        // ------------------------------------------------------------
        DataSet data = new DataSet();                          // empty dataset – no 'person' table or column

        // ------------------------------------------------------------
        // 4. Configure the ReportingEngine to allow missing members and
        //    specify a custom message that will be printed instead.
        // ------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers, // enable missing‑member handling
            MissingMemberMessage = "Missing"                 // custom placeholder text
        };

        // ------------------------------------------------------------
        // 5. Build the report using the template and the empty data source.
        // ------------------------------------------------------------
        engine.BuildReport(doc, data, "");                     // build report

        // ------------------------------------------------------------
        // 6. Save the resulting document.
        // ------------------------------------------------------------
        const string resultPath = "Result.docx";
        doc.Save(resultPath);                                 // save rule applied
    }
}
