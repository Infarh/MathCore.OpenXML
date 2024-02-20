using MathCore.OpenXML.ExcelProcessing.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting.Extensions;

namespace MathCore.OpenXML.Tests.ExcelProcessing.Extensions;

[TestClass]
public class ExcelCellExTests
{
    [TestMethod]
    public void GetCellRowIndexTest()
    {
        const int row_index_01 = 1;
        const int row_index_26 = 26;
        const int row_index_27 = 27;

        const string row_index_01_expected = "A";
        const string row_index_26_expected = "Z";
        const string row_index_27_expected = "AA";

        var row_index_01_actual = ExcelCellEx.GetCellRowIndex(row_index_01);
        var row_index_26_actual = ExcelCellEx.GetCellRowIndex(row_index_26);
        var row_index_27_actual = ExcelCellEx.GetCellRowIndex(row_index_27);

        row_index_01_actual.AssertEquals(row_index_01_expected);
        row_index_26_actual.AssertEquals(row_index_26_expected);
        row_index_27_actual.AssertEquals(row_index_27_expected);
    }

    [TestMethod]
    public void GetCellIndexTest()
    {
        const int row_index_01 = 1;
        const int row_index_26 = 26;
        const int row_index_27 = 27;

        const string row_index_01_expected = "A1";
        const string row_index_26_expected = "Z200";
        const string row_index_27_expected = "AA1000";

        var row_index_01_actual = ExcelCellEx.GetCellIndex(1, row_index_01);
        var row_index_26_actual = ExcelCellEx.GetCellIndex(200, row_index_26);
        var row_index_27_actual = ExcelCellEx.GetCellIndex(1000, row_index_27);

        row_index_01_actual.AssertEquals(row_index_01_expected);
        row_index_26_actual.AssertEquals(row_index_26_expected);
        row_index_27_actual.AssertEquals(row_index_27_expected);
    }
}
