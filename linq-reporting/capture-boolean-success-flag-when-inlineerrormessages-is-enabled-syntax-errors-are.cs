using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InlineErrorCapture
{
    // Custom evaluator that records the evaluation result and any error message.
    private class CaptureEvaluator : IComparisonExpressionEvaluator
    {
        public ComparisonEvaluationResult EvaluationResult { get; private set; }

        public CaptureEvaluator(ComparisonEvaluationResult result)
        {
            EvaluationResult = result;
        }

        public ComparisonEvaluationResult Evaluate(Field field, ComparisonExpression expression)
        {
            // Return the pre‑created result (could be success or error).
            return EvaluationResult;
        }
    }

    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable inline error messages for fields (if the option exists).
        // This makes the field display its error message directly in the document.
        // Note: The property may be called InlineErrorMessages in newer versions of Aspose.Words.
        // If the property does not exist, this line can be omitted or replaced with the appropriate API.
        // doc.FieldOptions.InlineErrorMessages = true; // Uncomment if supported.

        // Insert a COMPARE field with an intentional syntax error (e.g., missing right operand).
        FieldCompare field = (FieldCompare)builder.InsertField(FieldType.FieldCompare, true);
        field.LeftExpression = "5";
        field.ComparisonOperator = "=";
        field.RightExpression = ""; // Empty right expression causes a syntax error.

        // Prepare a failed evaluation result containing the error message.
        // The boolean success flag (Result) will be false, and ErrorMessage will hold details.
        ComparisonEvaluationResult errorResult = new ComparisonEvaluationResult("Syntax error: right expression is missing.");

        // Attach the custom evaluator to capture the result.
        doc.FieldOptions.ComparisonExpressionEvaluator = new CaptureEvaluator(errorResult);

        // Update fields – the custom evaluator will be invoked.
        doc.UpdateFields();

        // Retrieve the captured evaluation result.
        CaptureEvaluator evaluator = (CaptureEvaluator)doc.FieldOptions.ComparisonExpressionEvaluator;
        bool success = evaluator.EvaluationResult.Result;               // Expected to be false.
        string errorMessage = evaluator.EvaluationResult.ErrorMessage; // Contains the syntax error description.

        // Output the captured information (for demonstration purposes).
        Console.WriteLine($"Success flag: {success}");
        Console.WriteLine($"Error message: {errorMessage}");

        // Save the document (optional).
        doc.Save("InlineErrorCapture.docx");
    }
}
