using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working folder – the current directory.
        string workDir = Directory.GetCurrentDirectory();

        // Paths for the template, the XML data source and the generated report.
        string templatePath = Path.Combine(workDir, "Template.docx");
        string xmlPath = Path.Combine(workDir, "Orders.xml");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple XML file.  Numeric values are written with a dot
        //    as the decimal separator (invariant culture) so that the
        //    ReportingEngine can infer the correct numeric types.
        // -----------------------------------------------------------------
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
            <Orders>
                <Order>
                    <Id>1</Id>
                    <Amount>1234.56</Amount>
                </Order>
                <Order>
                    <Id>2</Id>
                    <Amount>7890.12</Amount>
                </Order>
            </Orders>";
        File.WriteAllText(xmlPath, xmlContent);

        // -----------------------------------------------------------------
        // 2. Build a Word template programmatically and embed LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln(); // empty line

        // foreach over the collection named "Order".
        builder.Writeln("<<foreach [order in Order]>>");
        builder.Writeln("Id: <<[order.Id]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template back (simulating a real‑world scenario where the
        //    template is a file).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Create an XmlDataSource from the XML file.
        // -----------------------------------------------------------------
        using (FileStream xmlStream = File.OpenRead(xmlPath))
        {
            XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

            // -----------------------------------------------------------------
            // 5. Build the report using ReportingEngine.
            //    Provide a data source name ("Order") so that the engine can
            //    resolve the collection correctly.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, xmlDataSource, "Order");
        }

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
