/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using System.IO;
using System.Xml;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ListDataValidationTests : ValidationTestBase
    {
        private IExcelDataValidationList _validation;

        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
            _validation = _sheet.Workbook.Worksheets[1].DataValidations.AddListValidation("A1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
        }

        [TestMethod]
        public void ListDataValidation_FormulaIsSet()
        {
            Assert.IsNotNull(_validation.Formula);
        }

        [TestMethod]
        public void ListDataValidation_CanAssignFormula()
        {
            _validation.Formula.ExcelFormula = "abc!A2";
            Assert.AreEqual("abc!A2", _validation.Formula.ExcelFormula);
        }
        [TestMethod]
        public void ListDataValidation_CanAssignDefinedName()
        {
            _validation.Formula.ExcelFormula = "ListData";
            Assert.AreEqual("ListData", _validation.Formula.ExcelFormula);
        }

        [TestMethod]
        public void ListDataValidation_WhenOneItemIsAddedCountIs1()
        {
            // Act
            _validation.Formula.Values.Add("test");

            // Assert
            Assert.AreEqual(1, _validation.Formula.Values.Count);
        }

        [TestMethod]
        public void ListDataValidation_ShouldNotThrowWhenNoFormulaOrValueIsSet()
        {
            _validation.Validate();
        }

        [TestMethod]
        public void ListDataValidation_ShowErrorMessageIsSet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list formula");

                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].Value = "A";
                sheet2.Cells["A2"].Value = "B";
                sheet2.Cells["A3"].Value = "C";

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("A1");
                // Alternatively:
                // var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
                validation.ShowErrorMessage = true;
                validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                validation.ErrorTitle = "An invalid value was entered";
                validation.Error = "Select a value from the list";
                validation.Formula.ExcelFormula = "Sheet2!A1:A3";

                Assert.IsTrue(validation.ShowErrorMessage.Value);
            }
        }

        [TestMethod]
        public void ListDataValidationExt_ShowDropDownIsSet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list formula");

                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].Value = "A";
                sheet2.Cells["A2"].Value = "B";
                sheet2.Cells["A3"].Value = "C";

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("A1");
                // Alternatively:
                // var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
                validation.HideDropDown = true;
                validation.ShowErrorMessage = true;
                validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                validation.ErrorTitle = "An invalid value was entered";
                validation.Error = "Select a value from the list";
                validation.Formula.ExcelFormula = "Sheet2!A1:A3";

                // refresh the data validation
                validation = sheet.DataValidations.Find(x => x.Uid == validation.Uid).As.ListValidation;

                Assert.IsTrue(validation.HideDropDown.Value);
                var v = validation as ExcelDataValidationList;
                var attributeValue = v.HideDropDown.Value;
                Assert.IsTrue(attributeValue);
            }
        }

        [TestMethod]
        public void ListDataValidation_ShowDropDownIsSet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list formula");
                sheet.Cells["A1"].Value = "A";
                sheet.Cells["A2"].Value = "B";
                sheet.Cells["A3"].Value = "C";

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("B1");
                validation.HideDropDown = true;
                validation.ShowErrorMessage = true;
                validation.Formula.ExcelFormula = "A1:A3";

                Assert.IsTrue(validation.HideDropDown.Value);
                var v = validation as ExcelDataValidationList;
                var attributeValue = v.HideDropDown.Value;
                Assert.IsTrue(attributeValue);
            }
        }

        [TestMethod]
        public void ListDataValidation_AllowsOperatorShouldBeFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("list operator");

                // add a validation and set values
                var validation = sheet.DataValidations.AddListValidation("A1");

                Assert.IsFalse(validation.AllowsOperator);
            }
        }

        [TestMethod]
        public void CompletelyEmptyListValidationsShouldNotThrow()
        {
            var excel = new ExcelPackage();
            var sheet = excel.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Column1";
            sheet.Cells["A2:A1048576"].DataValidation.AddListDataValidation();

            excel.Save();
        }

        [TestMethod]
        public void EmptyListValidationsShouldNotThrow()
        {
            var excel = new ExcelPackage();
            var sheet = excel.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Column1";
            var boolValidator = sheet.Cells["A2:A1048576"].DataValidation.AddListDataValidation();
            {
                boolValidator.Formula.Values.Add("");
                boolValidator.Formula.Values.Add("True");
                boolValidator.Formula.Values.Add("False");
                boolValidator.ShowErrorMessage = true;
                boolValidator.Error = "Choose either False or True";
            }

            excel.Save();
        }

        [TestMethod]
        public void ListLocalAndExt()
        {
            using (var package = OpenPackage("DataValidationExtLocalList.xlsx", true))
            {
                var ws1 = package.Workbook.Worksheets.Add("Worksheet1");
                package.Workbook.Worksheets.Add("Worksheet2");

                var localVal = ws1.DataValidations.AddDecimalValidation("A1:A5");

                localVal.Formula.Value = 0;
                localVal.Formula2.Value = 0.1;

                var extVal = ws1.DataValidations.AddListValidation("B1:B5");

                extVal.Formula.ExcelFormula = "Worksheet2!$G$12:G15";

                SaveAndCleanup(package);

                var p = OpenPackage("DataValidationExtLocalList.xlsx");

                var stream = new MemoryStream();
                p.SaveAs(stream);

                ExcelPackage pck = new ExcelPackage(stream);

                var stream2 = new MemoryStream();
                pck.SaveAs(stream2);
            }
        }

        //s664
        [TestMethod]
        public void BoolParsing()
        {
            using (var package = OpenTemplatePackage("s664.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[0];
                var cellVal = sheet.Cells["A1"].Value;
            }
        }
        //s664
        [TestMethod]
        public void ZeroShowDropDownShouldNotThrow()
        {
            //Test checks if showDropDown = 0 can be read.
            //We cannot create this state within epplus or excel normally.
            //So let's use a string.
            string ZeroDropDownData = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' mc:Ignorable='x14ac xr xr2 xr3' xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac' xmlns:xr='http://schemas.microsoft.com/office/spreadsheetml/2014/revision' xmlns:xr2='http://schemas.microsoft.com/office/spreadsheetml/2015/revision2' xmlns:xr3='http://schemas.microsoft.com/office/spreadsheetml/2016/revision3' xr:uid='{CA8FD7CC-E246-49DC-A925-78ADCCAB8017}'><dataValidations count='1'><dataValidation type='list' allowBlank='1' showDropDown='0' showInputMessage='1' showErrorMessage='1' sqref='A1:A9' xr:uid='{3B036BC1-69BC-45FF-B77B-4775A658E050}'><formula1>$D$1:$D$9</formula1></dataValidation></dataValidations></worksheet>";

            using (var dataReader = new StringReader(ZeroDropDownData))
            {
                var reader = XmlReader.Create(dataReader);
                reader.Read();
                reader.Read();
                reader.Read();

                using (var pck = OpenPackage("DataValidationListBoolWs.xlsx", true))
                {
                    var testSheet = pck.Workbook.Worksheets.Add("testSheet");
                    for(int i = 1; i< 9; i++)
                    {
                        testSheet.Cells[i, 4].Value = i;
                    }

                    var collection = new ExcelDataValidationCollection(reader, testSheet);
                    testSheet._dataValidations = collection;
                    SaveAndCleanup(pck);
                }
            }


            using (var pck = OpenPackage("DataValidationListBoolWs.xlsx"))
            {
                var testSheet = pck.Workbook.Worksheets.GetByName("testSheet");
                Assert.IsFalse(testSheet.DataValidations[0].As.ListValidation.HideDropDown);
                SaveAndCleanup(pck);
            }
        }

    }
}
