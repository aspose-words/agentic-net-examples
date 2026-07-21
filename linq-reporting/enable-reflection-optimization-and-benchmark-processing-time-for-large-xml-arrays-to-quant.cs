using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        const int itemCount = 5000;
        const string xmlPath = "data.xml";
        const string templatePath = "template.docx";
        const string outputWithoutOpt = "report_without_optimization.docx";
        const string outputWithOpt = "report_with_optimization.docx";

        // 1. Create large XML data file
        var sb = new StringBuilder();
        sb.AppendLine("<Orders>");
        for (int i = 1; i <= itemCount; i++)
        {
            sb.AppendLine("  <Order>");
            sb.AppendLine($"    <Id>{i}</Id>");
            sb.AppendLine($"    <Name>Order {i}</Name>");
            sb.AppendLine("  </Order>");
        }
        sb.AppendLine("</Orders>");
        File.WriteAllText(xmlPath, sb.ToString(), Encoding.UTF8);

        // 2. Create LINQ Reporting template
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Id: <<[order.Id]>>   Name: <<[order.Name]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 3. Load XML data source from the file content (use a stream to avoid path interpretation)
        string xmlContent = File.ReadAllText(xmlPath, Encoding.UTF8);
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));
        var xmlDataSource = new XmlDataSource(xmlStream);

        // 4. Benchmark without reflection optimization
        ReportingEngine.UseReflectionOptimization = false;
        var docWithoutOpt = new Document(templatePath);
        var engineWithoutOpt = new ReportingEngine();
        var stopwatch = Stopwatch.StartNew();
        engineWithoutOpt.BuildReport(docWithoutOpt, xmlDataSource, "Orders");
        stopwatch.Stop();
        docWithoutOpt.Save(outputWithoutOpt);
        long timeWithoutOpt = stopwatch.ElapsedMilliseconds;

        // 5. Benchmark with reflection optimization
        ReportingEngine.UseReflectionOptimization = true;
        var docWithOpt = new Document(templatePath);
        var engineWithOpt = new ReportingEngine();
        stopwatch.Restart();
        engineWithOpt.BuildReport(docWithOpt, xmlDataSource, "Orders");
        stopwatch.Stop();
        docWithOpt.Save(outputWithOpt);
        long timeWithOpt = stopwatch.ElapsedMilliseconds;

        // 6. Output results
        Console.WriteLine($"Processing {itemCount} items:");
        Console.WriteLine($"Without reflection optimization: {timeWithoutOpt} ms");
        Console.WriteLine($"With reflection optimization   : {timeWithOpt} ms");
    }
}
