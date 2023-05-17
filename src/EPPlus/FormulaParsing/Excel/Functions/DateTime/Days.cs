﻿/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Calculates the number of days between 2 dates",
        IntroducedInExcelVersion = "2013",
        SupportsArrays = true)]
    public class Days : ExcelFunction
    {
        private readonly ArrayBehaviourConfig _arrayConfig = new ArrayBehaviourConfig
        {
            ArrayParameterIndexes = new List<int> { 0, 1, 2 }
        };

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override ArrayBehaviourConfig GetArrayBehaviourConfig()
        {
            return _arrayConfig;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            ValidateArguments(arguments, 2);
            var numDate1 = ArgToDecimal(arguments, 0);
            var numDate2 = ArgToDecimal(arguments, 1);
            var endDate = System.DateTime.FromOADate(numDate1);
            var startDate = System.DateTime.FromOADate(numDate2);
            return CreateResult(endDate.Subtract(startDate).TotalDays, DataType.Date);
        }
        /// <summary>
        /// If the function has a namespace prefix when it's saved. Excel uses this for newer function. 
        /// For example "xlfn.".
        /// </summary>
        public override string NamespacePrefix => "xlfn.";
    }
}
