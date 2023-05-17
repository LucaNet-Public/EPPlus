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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Returns the year from a user-supplied date",
        SupportsArrays = true)]
    internal class Year : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var dateObj = arguments.ElementAt(0).Value;
            System.DateTime date;
            if (dateObj is string)
            {
                date = System.DateTime.Parse(dateObj.ToString());
            }
            else
            {
                var d = ArgToDecimal(arguments, 0);
                date = ConvertUtil.FromOADateExcel(d);
            }
            var aoDate = date.ToOADate();
            if (aoDate<0)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            else if(aoDate <= 1)
            {
                return CreateResult(1900, DataType.Integer);
            }
            return CreateResult(date.Year, DataType.Integer);
        }
    }
}
