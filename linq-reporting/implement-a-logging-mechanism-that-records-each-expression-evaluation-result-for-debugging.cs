using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Logger
{
    private static readonly List<string> _entries = new();

    public static void Log(string message) => _entries.Add($"{DateTime.Now:O} - {message}");

    public static void Save(string filePath)
    {
        File.WriteAllLines(filePath, _entries);
    }
}

public class ReportModel
{
    private string _customerName = string.Empty;
    private double _amount;

    public ReportModel(string customerName, double amount)
    {
        _customerName = customerName;
        _amount = amount;
    }

    public string CustomerName
    {
        get
        {
            Logger.Log($"CustomerName evaluated: {_customerName}");
            return _customerName;
        }
        set => _customerName = value;
    }

    public double Amount
    {
        get
        {
            Logger.Log($"Amount evaluated: {_amount}");
            return _amount;
        }
        set => _amount = value;
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for files.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";
        const string logPath = "EvaluationLog.txt";

        // 1. Create the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("Amount: <<[model.Amount]>>");
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare the data model.
        var model = new ReportModel("John Doe", 1234.56);

        // 4. Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        doc.Save(reportPath);

        // 6. Persist the evaluation log.
        Logger.Save(logPath);
    }
}
