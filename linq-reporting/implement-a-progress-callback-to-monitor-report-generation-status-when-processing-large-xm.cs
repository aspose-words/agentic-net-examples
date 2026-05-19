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

        // Create a simple template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template back for report generation.
        var reportDoc = new Document(templatePath);

        // Sample XML data representing a collection of orders.
        const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <Amount>123.45</Amount>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <Amount>678.90</Amount>
    </Order>
</Orders>";

        // Prepare the XML data source.
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));
        var xmlDataSource = new XmlDataSource(xmlStream);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlDataSource, "Orders");

        // Save the generated report with a progress callback.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };
        reportDoc.Save("ReportOutput.docx", saveOptions);
    }

    // Callback that receives document saving progress notifications.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime = DateTime.Now;
        private const double MaxDurationSeconds = 10.0; // safety limit

        public void Notify(DocumentSavingArgs args)
        {
            Console.WriteLine($"Saving progress: {args.EstimatedProgress:P2}");
            if ((DateTime.Now - _startTime).TotalSeconds > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"Saving aborted after exceeding {MaxDurationSeconds} seconds. Progress={args.EstimatedProgress}");
        }
    }
}
