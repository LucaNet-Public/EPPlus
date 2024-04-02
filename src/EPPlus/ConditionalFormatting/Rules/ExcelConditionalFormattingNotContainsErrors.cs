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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingNotContainsErrors : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingNotContainsErrors
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingNotContainsErrors(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.NotContainsErrors,
                address,
                priority,
                worksheet
                )
        {
            UpdateFormula();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingNotContainsErrors(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.NotContainsErrors,
                address,
                worksheet,
                xr)
        {
        }

        internal ExcelConditionalFormattingNotContainsErrors(ExcelConditionalFormattingNotContainsErrors copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
        }

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            var range = new ExcelRange(_ws, address.Address);
            return !ExcelErrorValue.Values.IsErrorValue(range.Value);
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingNotContainsErrors(this, newWs);
        }

        public override ExcelAddress Address 
        { 
            get { return base.Address; } 
            set { base.Address = value; UpdateFormula(); } 
        }

        void UpdateFormula()
        {
            if (Address != null)
            {
                Formula = string.Format(
                  "NOT(ISERROR({0}))",
                  Address.Start.Address);
            }
            else
            {
                Formula = string.Format(
                  "NOT(ISERROR({0}))",
                  "#REF!");
            }
        }

        #endregion Constructors

        /****************************************************************************************/
    }
}
