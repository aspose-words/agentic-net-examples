using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

namespace Aspose.Words.Tests
{
    // Simple integration test without external test frameworks.
    // The test is executed from Main and reports success/failure via console output.
    public class InlineErrorMessagesTests
    {
        // Custom evaluator that returns a predefined ComparisonEvaluationResult.
        private class ComparisonExpressionEvaluator : IComparisonExpressionEvaluator
        {
            private readonly ComparisonEvaluationResult _result;
            public readonly List<string[]> Invocations = new List<string[]>();

            public ComparisonExpressionEvaluator(ComparisonEvaluationResult result)
            {
                _result = result;
            }

            public ComparisonEvaluationResult Evaluate(Field field, ComparisonExpression expression)
            {
                // Record the arguments for possible further verification.
                Invocations.Add(new[]
                {
                    expression.LeftExpression,
                    expression.ComparisonOperator,
                    expression.RightExpression
                });

                return _result;
            }
        }

        private static void RunTest()
        {
            // Arrange: create a new document and builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an IF field with placeholder expressions.
            // The actual expressions are irrelevant because the evaluator will override the result.
            Field field = builder.InsertField(" IF \"left\" = \"right\" \"True\" \"False\" ", null);

            // Define an error message that should be embedded.
            const string expectedErrorMessage = "Comparison failed due to invalid data";

            // Create a ComparisonEvaluationResult that represents a failed evaluation.
            ComparisonEvaluationResult evalResult = new ComparisonEvaluationResult(expectedErrorMessage);

            // Attach the custom evaluator to the document.
            var evaluator = new ComparisonExpressionEvaluator(evalResult);
            doc.FieldOptions.ComparisonExpressionEvaluator = evaluator;

            // Act: update fields to trigger evaluation.
            doc.UpdateFields();

            // Assert: the evaluator's result contains the expected error message.
            if (evalResult.ErrorMessage != expectedErrorMessage)
                throw new Exception($"ErrorMessage mismatch. Expected: '{expectedErrorMessage}', Actual: '{evalResult.ErrorMessage}'");

            // The field's displayed result should be empty because the evaluation failed.
            if (!string.IsNullOrEmpty(field.Result))
                throw new Exception($"Field result should be empty on error, but was: '{field.Result}'");

            // Success flag: true if there is no error message.
            bool success = string.IsNullOrEmpty(evalResult.ErrorMessage);
            if (success)
                throw new Exception("Success flag should be false when an error message is present.");

            // Verify that the evaluator was invoked exactly once with the correct arguments.
            if (evaluator.Invocations.Count != 1)
                throw new Exception($"Evaluator should be invoked once, but was invoked {evaluator.Invocations.Count} times.");

            string[] args = evaluator.Invocations[0];
            if (args[0] != "left" || args[1] != "=" || args[2] != "right")
                throw new Exception($"Evaluator arguments mismatch. Got: [{args[0]}, {args[1]}, {args[2]}]");

            Console.WriteLine("Test passed.");
        }

        public static void Main()
        {
            try
            {
                RunTest();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Test failed: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
