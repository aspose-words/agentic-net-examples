using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create sample XML data source.
            // -----------------------------------------------------------------
            const string xmlFileName = "people.xml";
            File.WriteAllText(xmlFileName,
                @"<Persons>
                    <Person>
                        <Name>John Doe</Name>
                        <Age>30</Age>
                    </Person>
                    <Person>
                        <Name>Jane Smith</Name>
                        <Age>25</Age>
                    </Person>
                    <Person>
                        <Name>Bob Johnson</Name>
                        <Age>40</Age>
                    </Person>
                </Persons>");

            // -----------------------------------------------------------------
            // 2. Build a template document that contains a table with a foreach tag.
            // -----------------------------------------------------------------
            const string templateFileName = "template.docx";
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Persons Report");

            // Begin foreach loop over the XML collection named "persons".
            builder.Writeln("<<foreach [p in persons]>>");

            // Create a table that will be repeated for each person.
            Table table = builder.StartTable();

            // Table header.
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Age");
            builder.EndRow();

            // Row that will be repeated for each person.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Age]>>");
            builder.EndRow();

            // End of table and foreach.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            template.Save(templateFileName);

            // -----------------------------------------------------------------
            // 3. Load the template and generate the report using the XML data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templateFileName);

            using (FileStream xmlStream = File.OpenRead(xmlFileName))
            {
                XmlDataSource dataSource = new XmlDataSource(xmlStream);

                ReportingEngine engine = new ReportingEngine();
                engine.Options = ReportBuildOptions.None; // No special options required.
                engine.BuildReport(reportDoc, dataSource, "persons");
            }

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputFileName = "output.docx";
            reportDoc.Save(outputFileName);
        }
    }
}
