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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Returns the Arctangent of a given pair of x and y coordinates",
        SupportsArrays = true)]
    internal class Atan2 : ExcelFunction
    {
        private readonly ArrayBehaviourConfig _arrayConfig = new ArrayBehaviourConfig
        {
            ArrayParameterIndexes = new List<int> { 0, 1 }
        };

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override ArrayBehaviourConfig GetArrayBehaviourConfig()
        {
            return _arrayConfig;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arg1 = ArgToDecimal(arguments, 0);
            var arg2 = ArgToDecimal(arguments, 1);
            // Had to switch order of the arguments to get the same result as in excel /MA
            return CreateResult(System.Math.Atan2(arg2, arg1), DataType.Decimal);
        }
    }
}
