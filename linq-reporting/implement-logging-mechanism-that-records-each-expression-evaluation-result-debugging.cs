using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document that contains IF and COMPARE fields.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // IF field: 5 = 2 + 3  -> True
        builder.Writeln("IF field result: ");
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        ifField.LeftExpression = "5";
        ifField.ComparisonOperator = "=";
        ifField.RightExpression = "2 + 3";
        ifField.TrueText = "True";
        ifField.FalseText = "False";

        // COMPARE field: 3 < 2 -> False (0)
        builder.Writeln("COMPARE field result: ");
        FieldCompare compareField = (FieldCompare)builder.InsertField(FieldType.FieldCompare, true);
        compareField.LeftExpression = "3";
        compareField.ComparisonOperator = "<";
        compareField.RightExpression = "2";

        // -----------------------------------------------------------------
        // 2. Attach a custom evaluator that logs every comparison expression.
        // -----------------------------------------------------------------
        var evaluator = new ComparisonExpressionLogger();
        template.FieldOptions.ComparisonExpressionEvaluator = evaluator;

        // -----------------------------------------------------------------
        // 3. Build the report (no external data source is required).
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // BuildReport returns a bool indicating success; we ignore it here.
        engine.BuildReport(template, new object());

        // -----------------------------------------------------------------
        // 4. After the report is built, log the final results of each field.
        // -----------------------------------------------------------------
        foreach (Field field in template.Range.Fields)
        {
            if (field is FieldIf fi)
            {
                // EvaluateCondition returns an enum; we also log the displayed result.
                evaluator.LogResult($"{fi.GetFieldCode()} => Condition: {fi.EvaluateCondition()}, Result: \"{fi.Result}\"");
            }
            else if (field is FieldCompare fc)
            {
                // Result property already contains the evaluated 0/1 string.
                evaluator.LogResult($"{fc.GetFieldCode()} => Result: \"{fc.Result}\"");
            }
        }

        // -----------------------------------------------------------------
        // 5. Output the accumulated log.
        // -----------------------------------------------------------------
        Console.WriteLine("=== Expression Evaluation Log ===");
        Console.WriteLine(evaluator.GetLog());

        // -----------------------------------------------------------------
        // 6. Save the generated document (optional for visual verification).
        // -----------------------------------------------------------------
        template.Save("ReportWithEvaluationLog.docx");
    }
}

// ---------------------------------------------------------------------
// Custom evaluator that records each comparison expression and later
// allows arbitrary log entries (e.g., final field results).
// ---------------------------------------------------------------------
class ComparisonExpressionLogger : IComparisonExpressionEvaluator
{
    private readonly StringBuilder _log = new StringBuilder();

    // Called by Aspose.Words before each comparison is evaluated.
    public ComparisonEvaluationResult Evaluate(Field field, ComparisonExpression expression)
    {
        // Record the raw expression.
        _log.AppendLine($"Evaluating: {expression.LeftExpression} {expression.ComparisonOperator} {expression.RightExpression}");

        // Return null to let the default evaluation logic run.
        return null;
    }

    // Helper used by the main program to add additional log lines.
    public void LogResult(string message)
    {
        _log.AppendLine(message);
    }

    // Retrieve the full log as a string.
    public string GetLog()
    {
        return _log.ToString();
    }
}
