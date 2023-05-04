/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "5.1",
        Description = "Returns the Arcsine of a number",
        SupportsArrays = true)]
    internal class Asin : ExcelFunction
    {
        internal override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = ArgToDecimal(arguments, 0);
            return CreateResult(System.Math.Asin(arg), DataType.Decimal);
        }
    }
}
