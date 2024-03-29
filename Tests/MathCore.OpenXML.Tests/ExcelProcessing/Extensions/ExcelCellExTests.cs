﻿using MathCore.OpenXML.ExcelProcessing.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting.Extensions;

namespace MathCore.OpenXML.Tests.ExcelProcessing.Extensions;

[TestClass]
public class ExcelCellExTests
{
    [TestMethod]
    public void GetCellIndex()
    {
        const int row_index_01 = 1;
        const int row_index_26 = 26;
        const int row_index_27 = 27;

        const string row_index_01_expected = "A1";
        const string row_index_26_expected = "Z200";
        const string row_index_27_expected = "AA1000";

        var row_index_01_actual = ExcelCellEx.GetCellReference(1, row_index_01);
        var row_index_26_actual = ExcelCellEx.GetCellReference(200, row_index_26);
        var row_index_27_actual = ExcelCellEx.GetCellReference(1000, row_index_27);

        row_index_01_actual.AssertEquals(row_index_01_expected);
        row_index_26_actual.AssertEquals(row_index_26_expected);
        row_index_27_actual.AssertEquals(row_index_27_expected);
    }

    [TestMethod]
    public void GetCellRowIndex()
    {
        const string cell_reference_27 = "AA1000";
        const int expected_cel_index = 27;

        var actual_row_index = ExcelCellEx.GetCellRowIndex(cell_reference_27);

        actual_row_index.AssertEquals(expected_cel_index);
    }
}
