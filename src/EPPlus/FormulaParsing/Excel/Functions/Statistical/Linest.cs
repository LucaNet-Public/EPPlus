﻿/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/08/2023         EPPlus Software AB       Initial release EPPlus 7
*************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Sorting.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Statistical,
       EPPlusVersion = "7.0",
       Description = "The LINEST function calculates a regressional line that fits your data. It also calculates additional statistics." +
                     "It can handle several x-variables and perform multiple regression analysis.")]
    internal class Linest : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            //X can have more than one vector corresponding to each y-value
            if (!arguments[0].IsExcelRange) return CompileResult.GetErrorResult(eErrorType.Value);
            var constVar = true;
            var stats = false;
            bool multipleXranges = false;
            bool columnArray = false;
            bool rowArray = false;
            if (arguments.Count() > 1 && arguments[1].IsExcelRange)
            {
                var rangeX = arguments[1].ValueAsRangeInfo;
                var rangeY = arguments[0].ValueAsRangeInfo;
                var xColumns = rangeX.Size.NumberOfCols;
                var yColumns = rangeY.Size.NumberOfCols;
                var xRows = rangeX.Size.NumberOfRows;
                var yRows = rangeY.Size.NumberOfRows;

                if (arguments.Count() > 2 && arguments[2].DataType != DataType.Empty) constVar = ArgToBool(arguments, 2); //Need to change this
                if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);

                if ((xRows != yRows && xColumns == yColumns)
                    || (xColumns != yColumns && xRows == yRows))
                {
                    multipleXranges = true;
                }
                else
                {
                    if (xRows != yRows || xColumns != yColumns) return CreateResult(eErrorType.Ref);
                }

                RangeFlattener.GetNumericPairLists(rangeX, rangeY, !multipleXranges, out List<double> knownXsList, out List<double> knownYsList);

                //Converting the result of rangeflattener to double[]
                double[] knownXs = new double[knownXsList.Count];
                double[] knownYs = new double[knownYsList.Count];

                for (var i = 0; i < knownXsList.Count; i++)
                {
                    knownXs[i] = knownXsList[i];
                }
                
                for (var i = 0; i < knownYsList.Count; i++)
                {
                    knownYs[i] = knownYsList[i];
                }
                var r = 0;
                var c = 0;
                if (multipleXranges){
                    if (multipleXranges && xColumns != yColumns)
                    {
                        columnArray = true;

                        r = xRows;
                        c = xColumns;
                    }
                    else if (multipleXranges && xRows != yRows)
                    {
                        rowArray = true;
                        r = xColumns;
                        c = xRows;
                    }
                }
                else
                {
                    r = knownXs.Count();
                    c = 1;
                }
                if (multipleXranges && constVar)
                {
                    c += 1; //This is because we need to add a vector of ones to the matrix in order to force the multiple regression through (0, 0)
                }

                double[][] xRanges = MatrixHelper.CreateMatrix(r, c);
                //if (!multipleXranges)
                //{
                //    SortedDictionary<int, List<double>> xRanges = new SortedDictionary<int, List<double>>();
                //}
                //if (multipleXranges && xColumns != yColumns)
                //{
                //    columnArray = true;

                //    var colCount = -1;
                //    while (colCount < (xColumns - 1))
                //    {
                //        colCount += 1;
                //        var listCount = colCount;
                //        for(var i = 0; i < knownXs.Count(); i++)
                //        {
                //            xRanges[colCount][i] = knownXs[i];
                //            listCount += xColumns;
                //        }
                //    }
                //}
                if (columnArray)
                {
                    var counter = 0;
                    var delimiter = (constVar) ? xRanges[0].Count() - 1: xRanges[0].Count();
                    for (var i = 0; i < xRanges.Count(); i++)
                    {
                        for (var j = 0; j < delimiter; j++)
                        {
                            xRanges[i][j] = knownXs[counter];
                            counter += 1;
                        }
                    }
                }
                //if (multipleXranges && xColumns != yColumns)
                //{
                //    columnArray = true;
                //    for (var i = 0; i < xColumns; i++)
                //    {
                //        xRanges.Add(i, new List<double>());
                //    }

                //    var colCount = -1;

                //    while (colCount < (xColumns - 1))
                //    {
                //        colCount += 1;
                //        var listCount = colCount;
                //        while (listCount < knownXs.Count())
                //        {
                //            xRanges[colCount].Add(knownXs[listCount]);
                //            listCount += xColumns;
                //        }
                //    }
                //}
                //else if (multipleXranges && xRows != yRows)
                //{
                //    rowArray = true;
                //    for (var i = 0; i < xRows; i++)
                //    {
                //        xRanges.Add(i, new List<double>());
                //    }

                //    var rowCount = 0;
                //    var listCount = 0;
                //    while (rowCount < xRows)
                //    {
                //        rowCount += 1;
                //        while (listCount < (xColumns * rowCount))
                //        {
                //            xRanges[rowCount - 1].Add(knownXs[listCount]); //Write test for this
                //            listCount += 1;
                //        }
                //    }
                //}
                //else if (multipleXranges && xRows != yRows)
                //{
                //    rowArray = true;

                //    var rowCount = 0;
                //    var listCount = 0;
                //    while (rowCount < xRows)
                //    {
                //        rowCount += 1;
                //        while (listCount < (xColumns * rowCount))
                //        {
                //            xRanges[rowCount - 1][listCount] = knownXs[listCount];
                //            listCount += 1;
                //        }
                //    }
                //}
                else if (rowArray)
                {
                    //This shifts data thats row-based to column-based.
                    var counter = 0;
                    var delimiter = (constVar) ? xRanges[0].Count() - 1 : xRanges[0].Count();
                    for (var i = 0; i < delimiter; i++)
                    {
                        for (var j = 0; j < xRanges.Count(); j++)
                        {
                            xRanges[j][i] = knownXs[counter];
                            counter += 1;
                        }
                    }
                }

                //List<List<double>> xRangeList = new List<List<double>>();
                //double[][] xRangeList = new double[r][];
                //if (columnArray)
                //{
                //    for (var i = 0; i < xColumns; i++)
                //    {
                //        for (var j = 0; j < xRanges[i].Count(); j++)
                //        {
                //            xRangeList[i][j] = xRanges[i][j];
                //        }
                //    }
                //}
                //else if (rowArray)
                //{
                //    for (var i = 0; i < xRows; ++i)
                //    {
                //        for (var j = 0; j < xRanges[i].Count(); j++)
                //        {
                //            xRangeList[i][j] = xRanges[i][j];
                //        }
                //    }
                //}
                if (multipleXranges)
                {
                    var resultRangeX = LinestHelper.CalculateMultipleXRanges(knownYs, xRanges, constVar, stats);
                    return CreateResult(resultRangeX, DataType.ExcelRange);
                }
                else
                {
                    var resultRange = LinestHelper.CalculateResult(knownYs, knownXs, constVar, stats);
                    return CreateResult(resultRange, DataType.ExcelRange);
                }
            }
            else
            {
                var knownYsList = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[0] }, context, out ExcelErrorValue e1).ToList();
                if (e1 != null) return CreateResult(e1.Type);
                //knownYs = knownYs.Select(x => (double)x).ToList();
                var knownXsList = GetDefaultKnownXs(knownYsList.Count());

                //Working around jagged array issues
                double[] knownYs = new double[knownYsList.Count()];
                double[] knownXs = new double[knownXsList.Count()];

                for (var i = 0; i < knownYsList.Count(); i++)
                {
                    knownYs[i] = knownYsList[i];
                }

                for (var i = 0; i < knownXsList.Count(); i++)
                {
                    knownXs[i] = knownXsList[i];
                }

                if (arguments.Count() > 2) constVar = ArgToBool(arguments, 2);
                if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);

                var resultRange = LinestHelper.CalculateResult(knownYs, knownXs, constVar, stats); //change here so that multiple x is possible
                return CreateResult(resultRange, DataType.ExcelRange);
            }

        }
        private List<double> GetDefaultKnownXs(int count)
        {
            List<double> result = new List<double>();
            for (int i = 1; i <= count; i++)
            {
                result.Add(i);
            }
            return result;
        }
    }
}
