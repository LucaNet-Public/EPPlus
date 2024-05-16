﻿/*************************************************************************************************
Required Notice: Copyright (C) EPPlus Software AB. 
This software is licensed under PolyForm Noncommercial License 1.0.0 
and may only be used for noncommercial purposes 
https://polyformproject.org/licenses/noncommercial/1.0.0/

A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
Date               Author                       Change
*************************************************************************************************
01/18/2024         EPPlus Software AB       EPPlus 7.2
*************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfRunningTotal : PivotShowAsRunningTotal
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {
            var bf = fieldIndex.IndexOf(df.BaseField);
            var colFieldsStart = df.Field.PivotTable.RowFields.Count;
            var keyCol = fieldIndex.IndexOf(df.BaseField);
            var record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;
            var maxBfKey = 0;

            CalculateRunningTotal(df, fieldIndex, keys, ref calculatedItems, true);

            if (record.CacheItems[df.BaseField].Count(x => x is int) > 0)
            {
                maxBfKey = (int)record.CacheItems[df.BaseField].Where(x => x is int).Max();
            }

            foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
            {
                if (IsSumBefore(key.Key, bf, fieldIndex, colFieldsStart))
                {
                    calculatedItems[key.Key] = null;
                }
                else
                {
                    if (IsSumAfter(key.Key, bf, fieldIndex, colFieldsStart))
                    {
                        var parentKey = GetParentKey(key.Key, keyCol);
                        var parentValue = calculatedItems[parentKey];
                        if (parentValue is double pv)
                        {
                            if (calculatedItems[key.Key] is double v)
                            {
                                calculatedItems[key.Key] = (double)v / pv;
                            }
                        }
                        else if (calculatedItems[key.Key] is not ExcelErrorValue)
                        {
                            calculatedItems[key.Key] = parentValue;
                        }
                    }
                    else
                    {
                        calculatedItems[key.Key] = 1D;
                    }
                }
            }
        }
    }
}
