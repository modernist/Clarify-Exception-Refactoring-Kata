using System;
using System.Collections.Generic;
using System.Linq;

namespace codingdojo
{
    public class MessageEnricher
    {
        private static readonly List<(Predicate<Exception> Validate, Func<string, Exception, string> MessageFactory)> Validators;

        static MessageEnricher()
        {
            Validators = new List<(Predicate<Exception> Validator, Func<string, Exception, string> MessageFactory)>() {
                ((e) => e.Message.Equals("Object reference not set to an instance of an object") && e.StackTrace.Contains("VLookup"),
                (formulaName, e) => "Missing Lookup Table" ),

                ((e) => e.Message.Equals("Missing Formula") && e.GetType() == typeof(SpreadsheetException),
                (formulaName, e) => $"Invalid expression found in tax formula [{formulaName}]. Check for merged cells near {((SpreadsheetException)e).Cells}"),

                ((e) => e.Message.Equals("No matches found") && e.GetType() == typeof(SpreadsheetException),
                (formulaName, e) => $"No match found for token [{((SpreadsheetException)e).Token}] related to formula '{formulaName}'."),

                ((e) => e.GetType() == typeof(ExpressionParseException),
                (formulaName, e) => $"Invalid expression found in tax formula [{formulaName}]. Check that separators and delimiters use the English locale."),

                ((e) => e.Message.StartsWith("Circular Reference") && e.GetType() == typeof(SpreadsheetException),
                (formulaName, e) => $"Circular Reference in spreadsheet related to formula '{formulaName}'. Cells: {((SpreadsheetException)e).Cells}"),

                ((e) => true,
                (formulaName, e) => e.Message)
            };
        }

        public ErrorResult EnrichError(SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var formulaName = spreadsheetWorkbook.GetFormulaName();

            var error = Validators.First(ev => ev.Validate(e)).MessageFactory(formulaName, e);

            return new ErrorResult(formulaName, error, spreadsheetWorkbook.GetPresentation());
        }
    }
}