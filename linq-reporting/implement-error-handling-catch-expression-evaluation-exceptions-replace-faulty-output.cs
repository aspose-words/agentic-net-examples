using System;
using Aspose.Words;
using Aspose.Words.Fields;

class SafeComparisonEvaluator : IComparisonExpressionEvaluator
{
    // Evaluates a comparison expression and safely handles any errors.
    public ComparisonEvaluationResult Evaluate(Field field, ComparisonExpression expression)
    {
        try
        {
            // Attempt a simple numeric evaluation.
            double left = double.Parse(expression.LeftExpression);
            double right = double.Parse(expression.RightExpression);
            bool result;

            switch (expression.ComparisonOperator)
            {
                case "=":  result = left == right; break;
                case "<":  result = left <  right; break;
                case ">":  result = left >  right; break;
                case "<=": result = left <= right; break;
                case ">=": result = left >= right; break;
                case "<>": result = left != right; break;
                default:  throw new InvalidOperationException("Unsupported operator");
            }

            // Return a successful evaluation result.
            return new ComparisonEvaluationResult(result);
        }
        catch (Exception)
        {
            // On any exception, return a failed result with a placeholder message.
            return new ComparisonEvaluationResult("[Error]");
        }
    }
}

class Program
{
    static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an IF field that will cause an evaluation error (unsupported operator).
        builder.Writeln("Result of faulty IF field:");
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        ifField.LeftExpression = "10";
        ifField.ComparisonOperator = "/"; // unsupported operator triggers exception
        ifField.RightExpression = "0";
        ifField.TrueText = "True";
        ifField.FalseText = "False";

        // Assign the safe evaluator to handle errors gracefully.
        doc.FieldOptions.ComparisonExpressionEvaluator = new SafeComparisonEvaluator();

        // Update all fields – the faulty IF field will display the placeholder text.
        doc.UpdateFields();

        // Save the document (uses the provided save routine).
        doc.Save("FaultyIfField_Output.docx");
    }
}
