using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Reporting; // For XmlDataLoadOptions

namespace LinqReportingProgressDemo
{
    // Callback that receives document saving progress notifications.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            // Write the estimated progress percentage to the console.
            Console.WriteLine($"Saving progress: {args.EstimatedProgress:F2}%");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Required for some XML parsing scenarios.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // 1. Create a sample XML data file representing a large data set.
            const string xmlFileName = "data.xml";
            File.WriteAllText(xmlFileName,
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <Items>
            <Item>
                <Index>1</Index>
                <Name>Item A</Name>
            </Item>
            <Item>
                <Index>2</Index>
                <Name>Item B</Name>
            </Item>
        </Items>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <Items>
            <Item>
                <Index>1</Index>
                <Name>Item C</Name>
            </Item>
        </Items>
    </Order>
</Orders>");

            // 2. Build a Word template programmatically and embed LINQ Reporting tags.
            const string templateFileName = "template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Orders Report ===");
            builder.Writeln("<<foreach [order in Order]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items.Item]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templateFileName);

            // 3. Load the template and the XML data source.
            Document reportDoc = new Document(templateFileName);
            using (FileStream xmlStream = File.OpenRead(xmlFileName))
            {
                // Ensure the root object ("Orders") is always generated so that the engine can access its children.
                var loadOptions = new XmlDataLoadOptions { AlwaysGenerateRootObject = true };
                XmlDataSource xmlDataSource = new XmlDataSource(xmlStream, loadOptions);

                // 4. Configure the ReportingEngine.
                ReportingEngine engine = new ReportingEngine();
                engine.Options = ReportBuildOptions.None; // No special options are required for this simple example.

                // 5. Build the report. The data source name must match the root element name ("Orders").
                engine.BuildReport(reportDoc, xmlDataSource, "Orders");
            }

            // 6. Save the generated report while monitoring progress via the callback.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };
            const string outputFileName = "ReportWithProgress.docx";
            reportDoc.Save(outputFileName, saveOptions);

            // Indicate completion.
            Console.WriteLine($"Report generation completed. Output saved to '{outputFileName}'.");
        }
    }
}
