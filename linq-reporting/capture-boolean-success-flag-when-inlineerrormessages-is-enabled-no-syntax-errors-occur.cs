using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an IF field that will evaluate to true.
        // The field code: IF 1 = 1 "True" "False"
        Field field = builder.InsertField(" IF 1 = 1 \"True\" \"False\" ", null);

        // Create a comparison evaluation result that represents a successful evaluation.
        // The constructor with a bool creates a result with no error message.
        ComparisonEvaluationResult evalResult = new ComparisonEvaluationResult(true);

        // Capture the boolean success flag.
        bool success = evalResult.Result;          // true when the comparison succeeded
        string errorMessage = evalResult.ErrorMessage; // null because there is no error

        // Output the captured values (for demonstration purposes).
        Console.WriteLine($"Success flag: {success}");
        Console.WriteLine($"Error message: {(errorMessage ?? "none")}");

        // Assign a custom evaluator that always returns the above result.
        // This allows the field to use the same success flag during update.
        builder.Document.FieldOptions.ComparisonExpressionEvaluator = new FixedResultEvaluator(evalResult);

        // Update fields so the IF field uses the custom evaluator.
        doc.UpdateFields();

        // The field result will be "True" because the evaluation succeeded.
        Console.WriteLine($"Field result: {field.Result}");

        // Save the document if needed.
        // doc.Save("InlineErrorMessagesDemo.docx");
    }

    // Custom evaluator that returns a predefined ComparisonEvaluationResult.
    private class FixedResultEvaluator : IComparisonExpressionEvaluator
    {
        private readonly ComparisonEvaluationResult _result;

        public FixedResultEvaluator(ComparisonEvaluationResult result)
        {
            _result = result;
        }

        public ComparisonEvaluationResult Evaluate(Field field, ComparisonExpression expression)
        {
            // Return the predefined result; no error message means inline error messages are not needed.
            return _result;
        }
    }
}
