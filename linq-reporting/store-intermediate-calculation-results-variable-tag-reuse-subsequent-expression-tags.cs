using System;
using Aspose.Words;
using Aspose.Words.Fields;

class VariableReuseExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Insert a formula field that calculates an intermediate value.
        // ------------------------------------------------------------
        // Build the formula field: = 2 + 3
        FieldBuilder formulaBuilder = new FieldBuilder(FieldType.FieldFormula);
        formulaBuilder.AddArgument(2);
        formulaBuilder.AddArgument("+");
        formulaBuilder.AddArgument(3);
        FieldFormula formulaField = (FieldFormula)formulaBuilder.BuildAndInsert(builder.CurrentParagraph);
        // Force the field to evaluate so we can read its result.
        formulaField.Update();

        // ------------------------------------------------------------
        // 2. Store the result of the formula in a document variable.
        // ------------------------------------------------------------
        // The result of the formula field is available via the Result property.
        string intermediateResult = formulaField.Result.Trim();
        // Add (or replace) a variable named "SumResult" with the calculated value.
        doc.Variables.Add("SumResult", intermediateResult);

        // ------------------------------------------------------------
        // 3. Insert a DOCVARIABLE field that reuses the stored value.
        // ------------------------------------------------------------
        // This field will display the value of the variable "SumResult".
        FieldDocVariable variableField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        variableField.VariableName = "SumResult";
        variableField.Update();

        // ------------------------------------------------------------
        // 4. Demonstrate reuse in another expression (e.g., another formula).
        // ------------------------------------------------------------
        // Build a second formula that multiplies the stored variable by 4.
        // Use the variable field inside the formula by nesting it.
        FieldBuilder innerVariableBuilder = new FieldBuilder(FieldType.FieldDocVariable);
        innerVariableBuilder.AddArgument("SumResult"); // DOCVARIABLE SumResult
        FieldBuilder secondFormulaBuilder = new FieldBuilder(FieldType.FieldFormula);
        secondFormulaBuilder.AddArgument(innerVariableBuilder);
        secondFormulaBuilder.AddArgument("*");
        secondFormulaBuilder.AddArgument(4);
        FieldFormula secondFormula = (FieldFormula)secondFormulaBuilder.BuildAndInsert(builder.CurrentParagraph);
        secondFormula.Update();

        // ------------------------------------------------------------
        // 5. Save the document.
        // ------------------------------------------------------------
        doc.Save("VariableReuseExample.docx");
    }
}
