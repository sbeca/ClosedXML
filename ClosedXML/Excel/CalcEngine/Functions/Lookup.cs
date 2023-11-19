#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class Lookup
    {
        public static void Register(FunctionRegistry ce)
        {
            //ce.RegisterFunction("ADDRESS", , Address); // Returns a reference as text to a single cell in a worksheet
            //ce.RegisterFunction("AREAS", , Areas); // Returns the number of areas in a reference
            //ce.RegisterFunction("CHOOSE", , Choose); // Chooses a value from a list of values
            ce.RegisterFunction("COLUMN", 0, 1, Column, FunctionFlags.Range, AllowRange.All); // Returns the column number of a reference
            ce.RegisterFunction("COLUMNS", 1, 1, Adapt(Columns), FunctionFlags.Range, AllowRange.All); // Returns the number of columns in a reference
            //ce.RegisterFunction("FORMULATEXT", , Formulatext); // Returns the formula at the given reference as text
            //ce.RegisterFunction("GETPIVOTDATA", , Getpivotdata); // Returns data stored in a PivotTable report
            ce.RegisterFunction("HLOOKUP", 3, 4, AdaptLastOptional(Hlookup), FunctionFlags.Range, AllowRange.Only, 1); // Looks in the top row of an array and returns the value of the indicated cell
            ce.RegisterFunction("HYPERLINK", 1, 2, Adapt(Hyperlink), FunctionFlags.Scalar | FunctionFlags.SideEffect); // Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet
            ce.RegisterFunction("INDEX", 2, 4, Index, FunctionFlags.Range, AllowRange.Only, 0); // Uses an index to choose a value from a reference or array
            //ce.RegisterFunction("INDIRECT", , Indirect); // Returns a reference indicated by a text value
            //ce.RegisterFunction("LOOKUP", , Lookup); // Looks up values in a vector or array
            ce.RegisterFunction("MATCH", 2, 3, Match, FunctionFlags.Range, AllowRange.Only, 1); // Looks up values in a reference or array
            //ce.RegisterFunction("OFFSET", , Offset); // Returns a reference offset from a given reference
            ce.RegisterFunction("ROW", 0, 1, Row, FunctionFlags.Range, AllowRange.All); // Returns the row number of a reference
            ce.RegisterFunction("ROWS", 1, 1, Adapt(Rows), FunctionFlags.Range, AllowRange.All); // Returns the number of rows in a reference
            //ce.RegisterFunction("RTD", , Rtd); // Retrieves real-time data from a program that supports COM automation
            ce.RegisterFunction("TRANSPOSE", 1, 1, Adapt(Transpose), FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Returns the transpose of an array
            ce.RegisterFunction("VLOOKUP", 3, 4, AdaptLastOptional(Vlookup), FunctionFlags.Range, AllowRange.Only, 1); // Looks in the first column of an array and moves across the row to return the value of a cell
            ce.RegisterFunction("XLOOKUP", 3, 6, Xlookup, FunctionFlags.Range, AllowRange.Only, 1, 2);
        }

        private static AnyValue Column(CalcContext ctx, Span<AnyValue> p)
        {
            if (p.Length == 0 || p[0].IsBlank)
                return ctx.FormulaAddress.ColumnNumber;

            if (!p[0].TryPickArea(out var area, out var error))
                return error;

            var firstColumn = area.FirstAddress.ColumnNumber;
            var lastColumn = area.LastAddress.ColumnNumber;
            if (firstColumn == lastColumn)
                return firstColumn;

            var span = lastColumn - firstColumn + 1;
            var array = new ScalarValue[1, span];
            for (var col = firstColumn; col <= lastColumn; col++)
                array[0, col - firstColumn] = col;

            return new ConstArray(array);
        }

        private static AnyValue Columns(CalcContext _, AnyValue value)
        {
            return RowsOrColumns(value, false);
        }

        private static AnyValue Hlookup(CalcContext ctx, ScalarValue lookupValue, AnyValue rangeValue, ScalarValue rowIndex, ScalarValue flagValue)
        {
            if (lookupValue.IsError)
                return lookupValue.ToAnyValue();

            // Only the lookup value is converted to 0, not values in the range
            if (lookupValue.IsBlank)
                lookupValue = 0;

            if (lookupValue.TryPickText(out var lookupText, out _) && lookupText.Length > 255)
                return XLError.IncompatibleValue;

            if (rangeValue.TryPickScalar(out _, out var range))
                return XLError.NoValueAvailable;
            if (!range.TryPickT0(out var array, out var reference))
            {
                if (reference.Areas.Count > 1)
                    return XLError.NoValueAvailable;

                array = new ReferenceArray(reference.Areas.Single(), ctx);
            }

            if (!rowIndex.ToNumber(ctx.Culture).TryPickT0(out var row, out var error))
                return error;
            var rowIdx = (int)row;
            if (rowIdx < 1)
                return XLError.IncompatibleValue;
            if (rowIdx > array.Height)
                return XLError.CellReference;

            var approximateSearchFlag = true;
            if (!flagValue.IsBlank && !flagValue.TryCoerceLogicalOrBlankOrNumberOrText(out approximateSearchFlag, out var flagError))
                return flagError;

            if (approximateSearchFlag)
            {
                // Bisection in Excel and here differs, so we return different values for unsorted ranges, but same values for sorted ranges.
                var foundRow = Bisection(array, lookupValue);
                if (foundRow == -1)
                    return XLError.NoValueAvailable;

                return array[rowIdx - 1, foundRow].ToAnyValue();
            }
            else
            {
                // TODO: Implement wildcard search
                for (var columnIndex = 0; columnIndex < array.Width; columnIndex++)
                {
                    var currentValue = array[0, columnIndex];

                    // Because lookup value can't be an error, it doesn't matter that sort treats all errors as equal.
                    var comparison = ScalarValueComparer.SortIgnoreCase.Compare(currentValue, lookupValue);
                    if (comparison == 0)
                        return array[rowIdx - 1, columnIndex].ToAnyValue();
                }

                return XLError.NoValueAvailable;
            }
        }

        private static AnyValue Hyperlink(CalcContext ctx, string linkLocation, ScalarValue? friendlyName)
        {
            return friendlyName?.ToAnyValue() ?? linkLocation;
        }

        private static AnyValue Index(CalcContext ctx, Span<AnyValue> p)
        {
            var rangeValue = p[0];
            if (rangeValue.TryPickScalar(out _, out var range))
                return XLError.NoValueAvailable;
            if (!range.TryPickT0(out var rangeArray, out var rangeReference))
            {
                if (rangeReference.Areas.Count > 1)
                    return XLError.NoValueAvailable;

                rangeArray = new ReferenceArray(rangeReference.Areas.Single(), ctx);
            }

            if (!p[1].TryPickScalar(out var rowNumber, out var rowNumberAsCollection))
            {
                if (rowNumberAsCollection.TryPickT0(out var rowNumberAsArray, out var rowNumberAsReference))
                {
                    if (rowNumberAsArray.Width * rowNumberAsArray.Height == 1)
                        rowNumber = rowNumberAsArray[0, 0];
                    else
                        return XLError.IncompatibleValue;
                }
                else
                {
                    if (!rowNumberAsReference.TryGetSingleCellValue(out rowNumber, ctx))
                        return XLError.IncompatibleValue;
                }
            }

            ScalarValue columnNumber = 1;
            if (p.Length > 2)
            {
                if (!p[2].TryPickScalar(out columnNumber, out var columnNumberAsCollection))
                {
                    if (columnNumberAsCollection.TryPickT0(out var columnNumberAsArray, out var columnNumberAsReference))
                    {
                        if (columnNumberAsArray.Width * columnNumberAsArray.Height == 1)
                            columnNumber = columnNumberAsArray[0, 0];
                        else
                            return XLError.IncompatibleValue;
                    }
                    else
                    {
                        if (!columnNumberAsReference.TryGetSingleCellValue(out columnNumber, ctx))
                            return XLError.IncompatibleValue;
                    }
                }
            }

            if (rangeArray.Width > 1 && rangeArray.Height > 1)
            {
                var row_num = (int)rowNumber.GetNumber();
                var column_num = (int)columnNumber.GetNumber();

                if (row_num > rangeArray.Height)
                    return XLError.CellReference;

                if (column_num > rangeArray.Width)
                    return XLError.CellReference;

                return rangeArray[row_num - 1, column_num > 0 ? column_num - 1 : 0].ToAnyValue();
            }
            else if (p.Length == 2)
            {
                var cell_num = (int)rowNumber.GetNumber();
                if (rangeArray.Width > 1)
                {
                    if (cell_num > rangeArray.Width)
                        return XLError.CellReference;
                    else
                        return rangeArray[0, cell_num - 1].ToAnyValue();
                }
                else
                {
                    if (cell_num > rangeArray.Height)
                        return XLError.CellReference;
                    else
                        return rangeArray[cell_num - 1, 0].ToAnyValue();
                }
            }
            else
            {
                int column_num = 1;
                int row_num = 1;

                if (rowNumber.IsNumber)
                    row_num = (int)rowNumber.GetNumber();

                if (columnNumber.IsNumber)
                    column_num = (int)columnNumber.GetNumber();

                var rangeIsRow = rangeArray.Height == 1;
                if (rangeIsRow && row_num > 1)
                    return XLError.CellReference;

                if (!rangeIsRow && column_num > 1)
                    return XLError.CellReference;

                if (row_num > rangeArray.Height)
                    return XLError.CellReference;

                if (column_num > rangeArray.Width)
                    return XLError.CellReference;

                return rangeArray[row_num - 1, column_num > 0 ? column_num - 1 : 0].ToAnyValue();
            }
        }

        private static AnyValue Match(CalcContext ctx, Span<AnyValue> p)
        {
            if (!p[0].TryPickScalar(out var lookupValue, out var lookupValueAsCollection))
            {
                if (lookupValueAsCollection.TryPickT0(out var lookupValueAsArray, out var lookupValueAsReference))
                {
                    if (lookupValueAsArray.Width * lookupValueAsArray.Height == 1)
                        lookupValue = lookupValueAsArray[0, 0];
                    else
                        return XLError.IncompatibleValue;
                }
                else
                {
                    if (!lookupValueAsReference.TryGetSingleCellValue(out lookupValue, ctx))
                        return XLError.IncompatibleValue;
                }
            }

            var rangeValue = p[1];
            if (rangeValue.TryPickScalar(out _, out var range))
                return XLError.NoValueAvailable;
            if (!range.TryPickT0(out var rangeArray, out var rangeReference))
            {
                if (rangeReference.Areas.Count > 1)
                    return XLError.NoValueAvailable;

                var singleAreaRange = ctx.Worksheet.Range(rangeReference.Areas.Single());
                // Reduce the amount of work we have to do by excluding any unused cells from the end of the range
                var rangeAddressToUse = new XLRangeAddress((XLAddress)singleAreaRange.FirstCell().Address, (XLAddress)singleAreaRange.LastCellUsed().Address);
                rangeArray = new ReferenceArray(rangeAddressToUse, ctx);
            }

            int matchType = 1;
            if (p.Length > 2)
            {
                if (!p[2].TryPickScalar(out var matchTypeScalar, out _) || !matchTypeScalar.IsNumber)
                    return XLError.IncompatibleValue;

                matchType = Math.Sign(matchTypeScalar.GetNumber());
            }

            if (rangeArray.Width != 1 && rangeArray.Height != 1)
                return XLError.NoValueAvailable;

            Predicate<int> lookupPredicate = null;
            switch (matchType)
            {
                case 0:
                    lookupPredicate = i => i == 0;
                    break;

                case 1:
                    lookupPredicate = i => i <= 0;
                    break;

                case -1:
                    lookupPredicate = i => i >= 0;
                    break;

                default:
                    return XLError.NoValueAvailable;
            }

            Func<ScalarValue, ScalarValue, bool> comparePredicate = (c1, c2) =>
            {
                var compare = AnyValue.CompareValues(c1, c2, ctx.Culture);
                if (compare.TryPickT0(out var compareValue, out _))
                    return lookupPredicate.Invoke(compareValue);
                else
                    return false;
            };

            int foundValue = 0;
            var isFirst = true;
            var doneSearching = false;
            ScalarValue previousValue = ScalarValue.Blank;

            for (int y = 0; y < rangeArray.Height; ++y)
            {
                for (int x = 0; x < rangeArray.Width; ++x)
                {
                    var current = rangeArray[y, x];

                    if (matchType == 0)
                    {
                        if (comparePredicate.Invoke(current, lookupValue))
                            return (y * rangeArray.Width) + x + 1;
                    }
                    else
                    {
                        if (!isFirst)
                        {
                            // When matchType != 0, we have to assume that the order of the items being search is ascending or descending
                            if (!comparePredicate(previousValue, current))
                            {
                                doneSearching = true;
                                break;
                            }
                        }

                        isFirst = false;
                        previousValue = current;

                        if (comparePredicate(current, lookupValue))
                        {
                            foundValue = (y * rangeArray.Width) + x + 1;
                        }
                    }
                }

                if (doneSearching)
                    break;
            }

            if (foundValue > 0)
                return foundValue;
            else
                return XLError.NoValueAvailable;
        }

        private static AnyValue Row(CalcContext ctx, Span<AnyValue> p)
        {
            if (p.Length == 0 || p[0].IsBlank)
                return ctx.FormulaAddress.RowNumber;

            if (!p[0].TryPickArea(out var area, out var error))
                return error;

            var firstRow = area.FirstAddress.RowNumber;
            var lastRow = area.LastAddress.RowNumber;
            if (firstRow == lastRow)
                return firstRow;

            var span = lastRow - firstRow + 1;
            var array = new ScalarValue[span, 1];
            for (var row = firstRow; row <= lastRow; row++)
                array[row - firstRow, 0] = row;

            return new ConstArray(array);
        }

        private static AnyValue Rows(CalcContext _, AnyValue value)
        {
            return RowsOrColumns(value, true);
        }

        private static AnyValue Transpose(CalcContext ctx, AnyValue value)
        {
            if (value.TryPickSingleOrMultiValue(out var single, out var multi, ctx))
                return single.ToAnyValue();

            return new TransposedArray(multi);
        }

        private static AnyValue Vlookup(CalcContext ctx, ScalarValue lookupValue, AnyValue rangeValue, ScalarValue columnIndex, ScalarValue flagValue)
        {
            if (lookupValue.IsError)
                return lookupValue.ToAnyValue();

            // Only the lookup value is converted to 0, not values in the range
            if (lookupValue.IsBlank)
                lookupValue = 0;

            if (lookupValue.TryPickText(out var lookupText, out _) && lookupText.Length > 255)
                return XLError.IncompatibleValue;

            if (rangeValue.TryPickScalar(out _, out var range))
                return XLError.NoValueAvailable;
            if (!range.TryPickT0(out var array, out var reference))
            {
                if (reference.Areas.Count > 1)
                    return XLError.NoValueAvailable;

                array = new ReferenceArray(reference.Areas.Single(), ctx);
            }

            if (!columnIndex.ToNumber(ctx.Culture).TryPickT0(out var column, out var error))
                return error;
            var columnIdx = (int)column;
            if (columnIdx < 1)
                return XLError.IncompatibleValue;
            if (columnIdx > array.Width)
                return XLError.CellReference;

            var approximateSearchFlag = true;
            if (!flagValue.IsBlank && !flagValue.TryCoerceLogicalOrBlankOrNumberOrText(out approximateSearchFlag, out var flagError))
                return flagError;

            if (approximateSearchFlag)
            {
                // Bisection in Excel and here differs, so we return different values for unsorted ranges, but same values for sorted ranges.
                var foundRow = Bisection(array, lookupValue);
                if (foundRow == -1)
                    return XLError.NoValueAvailable;

                return array[foundRow, columnIdx - 1].ToAnyValue();
            }
            else
            {
                // TODO: Implement wildcard search
                for (var rowIndex = 0; rowIndex < array.Height; rowIndex++)
                {
                    var currentValue = array[rowIndex, 0];

                    // Because lookup value can't be an error, it doesn't matter that sort treats all errors as equal.
                    var comparison = ScalarValueComparer.SortIgnoreCase.Compare(currentValue, lookupValue);
                    if (comparison == 0)
                        return array[rowIndex, columnIdx - 1].ToAnyValue();
                }

                return XLError.NoValueAvailable;
            }
        }

        private static AnyValue Xlookup(CalcContext ctx, Span<AnyValue> p)
        {
            if (!p[0].TryPickScalar(out var lookupValue, out var lookupValueAsCollection))
            {
                if (!lookupValueAsCollection.TryPickT0(out _, out var lookupValueAsReference))
                {
                    if (!lookupValueAsReference.TryGetSingleCellValue(out lookupValue, ctx))
                        return XLError.IncompatibleValue;
                }
                else
                {
                    return XLError.IncompatibleValue;
                }
            }

            if (lookupValue.IsError)
                return lookupValue.ToAnyValue();

            var lookupRangeValue = p[1];
            var returnRangeValue = p[2];

            if (lookupValue.TryPickText(out var lookupText, out _) && lookupText.Length > 255)
                return XLError.IncompatibleValue;

            if (lookupRangeValue.TryPickScalar(out var lookupScalar, out var lookupRange))
            {
                if (ScalarValueComparer.SortIgnoreCase.Compare(lookupValue, lookupScalar) == 0)
                    return returnRangeValue;
                else
                    return XLError.NoValueAvailable;
            }

            if (!lookupRange.TryPickT0(out var lookupArray, out var lookupReference))
            {
                // Range must be contiguous
                if (lookupReference.Areas.Count > 1)
                    return XLError.NoValueAvailable;
                // Ranges are only allowed to be 1-dimensional
                if (lookupReference.Areas.Count > 0 && lookupReference.Areas[0].RowSpan > 1 && lookupReference.Areas[0].ColumnSpan > 1)
                    return XLError.IncompatibleValue;

                lookupArray = new ReferenceArray(lookupReference.Areas.Single(), ctx);
            }

            if (returnRangeValue.TryPickScalar(out _, out var returnRange))
                return XLError.IncompatibleValue;

            IXLRange returnReferenceRange = null;
            if (!returnRange.TryPickT0(out Array returnArray, out var returnReference))
            {
                // Range must be contiguous
                if (returnReference.Areas.Count > 1)
                    return XLError.NoValueAvailable;
                // Ranges are only allowed to be 1-dimensional
                if (returnReference.Areas.Count > 0 && returnReference.Areas[0].RowSpan > 1 && returnReference.Areas[0].ColumnSpan > 1)
                    return XLError.IncompatibleValue;

                var returnArraySingle = returnReference.Areas.Single();
                returnArray = new ReferenceArray(returnArraySingle, ctx);
                returnReferenceRange = (returnArraySingle.Worksheet ?? ctx.Worksheet).Range(returnArraySingle);
            }

            // The lengths of both ranges must be exactly the same
            if ((lookupArray.Width * lookupArray.Height) != (returnArray.Width * returnArray.Height))
                return XLError.IncompatibleValue;

            var ifNotFoundValue = ScalarValue.Blank;
            if (p.Length > 3 && !p[3].TryPickScalar(out ifNotFoundValue, out _))
                return XLError.IncompatibleValue;

            int matchModeInt = 0; // Default value
            if (p.Length > 4)
            {
                if (!p[4].TryPickScalar(out var matchModeValue, out _))
                    return XLError.IncompatibleValue;
                if (!matchModeValue.ToNumber(ctx.Culture).TryPickT0(out var matchMode, out var error))
                    return error;
                matchModeInt = (int)matchMode;
                if (matchModeInt < -1 || matchModeInt > 2)
                    return XLError.IncompatibleValue;
            }

            int searchModeInt = 0; // Default value
            if (p.Length > 5)
            {
                if (!p[5].TryPickScalar(out var searchModeValue, out _))
                    return XLError.IncompatibleValue;
                if (!searchModeValue.ToNumber(ctx.Culture).TryPickT0(out var searchMode, out var error))
                    return error;
                searchModeInt = (int)searchMode;
                if (searchModeInt < -1 || searchModeInt > 2)
                    return XLError.IncompatibleValue;
            }

            if (matchModeInt == 0) // 0 - Try to find exact match. If none found, return #N/A
            {
                if (lookupArray.Height > lookupArray.Width)
                {
                    for (var rowIndex = 0; rowIndex < lookupArray.Height; rowIndex++)
                    {
                        var currentValue = lookupArray[rowIndex, 0];

                        // Because lookup value can't be an error, it doesn't matter that sort treats all errors as equal.
                        var comparison = ScalarValueComparer.SortIgnoreCase.Compare(currentValue, lookupValue);
                        if (comparison == 0)
                            return (returnReferenceRange != null ? new Reference((XLRangeAddress)returnReferenceRange.Cell(rowIndex + 1, 1).AsRange().RangeAddress) : returnArray[rowIndex, 0].ToAnyValue());
                    }
                }
                else
                {
                    for (var columnIndex = 0; columnIndex < lookupArray.Width; columnIndex++)
                    {
                        var currentValue = lookupArray[0, columnIndex];

                        // Because lookup value can't be an error, it doesn't matter that sort treats all errors as equal.
                        var comparison = ScalarValueComparer.SortIgnoreCase.Compare(currentValue, lookupValue);
                        if (comparison == 0)
                            return (returnReferenceRange != null ? new Reference((XLRangeAddress)returnReferenceRange.Cell(1, columnIndex + 1).AsRange().RangeAddress) : returnArray[0, columnIndex].ToAnyValue());
                    }
                }

                return ifNotFoundValue.IsBlank ? XLError.NoValueAvailable : ifNotFoundValue.ToAnyValue();
            }
            else if (matchModeInt == -1) // -1 - Try to find exact match. If none found, return the next smaller item.
            {
                // Bisection in Excel and here differs, so we return different values for unsorted ranges, but same values for sorted ranges.
                var foundRow = Bisection(lookupArray, lookupValue);
                if (foundRow == -1)
                    return ifNotFoundValue.IsBlank ? XLError.NoValueAvailable : ifNotFoundValue.ToAnyValue();

                return (returnReferenceRange != null ? new Reference((XLRangeAddress)returnReferenceRange.Cell(foundRow + 1, 1).AsRange().RangeAddress) : returnArray[foundRow, 0].ToAnyValue());
            }
            else if (matchModeInt == 1) // 1 - Try to find exact match. If none found, return the next larger item.
            {
                // Bisection in Excel and here differs, so we return different values for unsorted ranges, but same values for sorted ranges.
                var foundRow = Bisection(lookupArray, lookupValue, true);
                if (foundRow == -1)
                    return ifNotFoundValue.IsBlank ? XLError.NoValueAvailable : ifNotFoundValue.ToAnyValue();

                return (returnReferenceRange != null ? new Reference((XLRangeAddress)returnReferenceRange.Cell(foundRow + 1, 1).AsRange().RangeAddress) : returnArray[foundRow, 0].ToAnyValue());
            }
            else if (matchModeInt == 2) // 2 - A wildcard match where *, ?, and ~ have special meaning: https://support.microsoft.com/en-us/office/using-wildcard-characters-in-searches-ef94362e-9999-4350-ad74-4d2371110adb
            {
                // TODO: Implement wildcard search
                return XLError.IncompatibleValue;
            }
            else
            {
                return XLError.IncompatibleValue;
            }
        }

        private static int Bisection(Array range, ScalarValue lookupValue, bool returnClosestMatchAboveLookupValue = false)
        {
            // Bisection is predicated on a fact that values of the same type are sorted.
            // If they are not, results are unpredictable.
            // Invariants:
            // * Low row has a value that is less or equal than lookup value
            // * High row has a value that is greater than lookup value
            var lowRow = 0;
            var highRow = range.Height - 1;

            lowRow = FindSameTypeRow(range, highRow, 1, lowRow, in lookupValue);
            if (lowRow == -1)
                return -1; // Range doesn't contain even one element of same type

            // Sanity check for unsorted ranges. For bisection to work, lowRow always
            // has to have a value that is less or equal to the lookup value.
            var lowValue = range[lowRow, 0];
            var lowCompare = ScalarValueComparer.SortIgnoreCase.Compare(lowValue, lookupValue);

            // Ensure invariants before main loop. If even lowest value in the range is greater than lookup value,
            // then there can't be any row that matches lookup value/lower.
            if (lowCompare > 0)
                return returnClosestMatchAboveLookupValue ? lowRow : -1;

            // Since we already know that there is at least one element of same type as lookup value,
            // high row will find something, though it might be same row as lowRow.
            highRow = FindSameTypeRow(range, lowRow, -1, highRow, in lookupValue);

            // Sanity check for unsorted ranges. For bisection to work, highRow always
            // has to have a value that is greater than the lookup value
            var highValue = range[highRow, 0];
            var highCompare = ScalarValueComparer.SortIgnoreCase.Compare(highValue, lookupValue);

            // Ensure invariants before main loop. If the lookup value is greater/equal than
            // the greatest value of the range, it is the result.
            if (highCompare <= 0)
                return returnClosestMatchAboveLookupValue ? -1 : highRow;

            // Now we have two borders with actual values and we know the lookup value is less than high and greater/equal to lower
            while (true)
            {
                // The FindMiddle method returns only values [lowRow, highRow)
                // so in each loop it decreases the interval. The lowRow value is
                // the last one checked during search of a middle.
                var middleRow = FindMiddle(range, lowRow, highRow, in lookupValue);

                // A condition for "if an exact match is not found, the next
                // largest value that is less than lookup-value is returned".
                // At this time, lowRow is less than lookup value and highRow
                // is more than lookup value.
                if (middleRow == lowRow)
                {
                    if (returnClosestMatchAboveLookupValue)
                    {
                        var lastMiddleValue = range[middleRow, 0];
                        if (ScalarValueComparer.SortIgnoreCase.Compare(lastMiddleValue, lookupValue) == 0)
                            return lowRow;
                        else
                            return highRow;
                    }
                    else
                    {
                        return lowRow;
                    }
                }

                var middleValue = range[middleRow, 0];
                var middleCompare = ScalarValueComparer.SortIgnoreCase.Compare(middleValue, lookupValue);

                if (middleCompare <= 0)
                    lowRow = middleRow;
                else
                    highRow = middleRow;
            }
        }

        /// <summary>
        /// Find a row with a value of same type as <paramref name="lookupValue"/>
        /// between values <paramref name="low"/> and <c><paramref name="high"/> - 1</c>.
        /// We know that both <paramref name="low"/> and <paramref name="high"/>
        /// contain value of the same type, so we always get a valid row.
        /// </summary>
        private static int FindMiddle(Array range, int low, int high, in ScalarValue lookupValue)
        {
            Debug.Assert(low < high);
            var middleRow = (low + high) / 2;

            // Since low is < high, it's always possible skip high row for determining middle row
            var higherIndex = FindSameTypeRow(range, high - 1, 1, middleRow, in lookupValue);
            if (higherIndex != -1)
                return higherIndex;

            // We can't skip low like we did for high, because there might be only different type
            // Cells between low row and high row.
            var lowerIndex = FindSameTypeRow(range, low, -1, middleRow, in lookupValue);
            return lowerIndex;
        }

        /// <summary>
        /// Find row index of an element with same type as the lookup value. Go from
        /// <paramref name="startRow"/> to the <paramref name="limitRow"/> by a step
        /// of <paramref name="delta"/>. If there isn't any such row, return <c>-1</c>.
        /// </summary>
        private static int FindSameTypeRow(Array range, int limitRow, int delta, int startRow, in ScalarValue lookupValue)
        {
            // Although the spec says that elements must be sorted in
            // "ascending order", as follows: ..., -2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE.
            // In reality, comparison ignores elements of the different type than lookupValue.
            // E.g. search for 2.5 in the {"1", 2, "3", #DIV/0!, 3 } will find the second element 2
            // Elements with incompatible type are just skipped.
            int currentRow;
            for (currentRow = startRow; !lookupValue.HaveSameType(range[currentRow, 0]); currentRow += delta)
            {
                // Don't move beyond limitRow
                if (currentRow == limitRow)
                    return -1;
            }

            return currentRow;
        }

        private static AnyValue RowsOrColumns(AnyValue value, bool rows)
        {
            if (value.TryPickArea(out var area, out _))
                return rows ? area.RowSpan : area.ColumnSpan;

            if (value.TryPickArray(out var array))
                return rows ? array.Height : array.Width;

            if (value.TryPickError(out var error))
                return error;

            if (value.IsLogical || value.IsNumber || value.IsText)
                return 1;

            if (value.IsBlank)
                return XLError.IncompatibleValue;

            // Only thing left, if reference has multiple areas
            return XLError.CellReference;
        }
    }
}
