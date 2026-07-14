using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a large XML data set.
        string xmlPath = Path.Combine(workDir, "data.xml");
        CreateLargeXml(xmlPath, 1000); // 1000 records.

        // 2. Create a LINQ Reporting template.
        string templatePath = Path.Combine(workDir, "template.docx");
        CreateTemplate(templatePath);

        // 3. Load the template document.
        Document templateDoc = new Document(templatePath);

        // 4. Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Use default options.
        // The root name in the template is "Items".
        engine.BuildReport(templateDoc, dataSource, "Items");

        // 6. Save the generated report with a progress callback.
        string reportPath = Path.Combine(workDir, "report.docx");
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };
        templateDoc.Save(reportPath, saveOptions);
    }

    // Generates an XML file with the specified number of <Item> elements.
    private static void CreateLargeXml(string filePath, int count)
    {
        using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
        {
            writer.WriteLine("<Items>");
            for (int i = 1; i <= count; i++)
            {
                writer.WriteLine("  <Item>");
                writer.WriteLine($"    <Id>{i}</Id>");
                writer.WriteLine($"    <Name>Item {i}</Name>");
                writer.WriteLine("  </Item>");
            }
            writer.WriteLine("</Items>");
        }
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Report of Items");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Callback that receives progress notifications while saving a document.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            // Output the estimated progress percentage to the console.
            Console.WriteLine($"Saving progress: {args.EstimatedProgress:F2}%");
        }
    }
}
