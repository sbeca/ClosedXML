using System;
using System.Runtime.CompilerServices;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLCellsUsedOptions
    {

        None                    = 0,
        NoConstraints           = None,

        Contents                = 1 << 0,
        DataType                = 1 << 1,
        NormalFormats           = 1 << 2,
        ConditionalFormats      = 1 << 3,
        Comments                = 1 << 4,
        DataValidation          = 1 << 5,
        MergedRanges            = 1 << 6,
        Sparklines              = 1 << 7,

        AllFormats = NormalFormats | ConditionalFormats,
        AllContents = Contents | DataType | Comments,
        All = Contents | DataType | NormalFormats | ConditionalFormats | Comments | DataValidation | MergedRanges | Sparklines
    }

    internal static class XLCellsUsedOptionsExtensions
    {
        /// <summary>
        /// Determines whether one or more bit fields are set in the current instance.
        /// </summary>
        /// <remarks>
        /// This is functionally the same as <see cref="Enum.HasFlag(Enum)"/>, just without the boxing
        /// and extra error checking required of a generic solution. This is to improve performance.
        /// </remarks>
        /// <param name="self">An enumeration value.</param>
        /// <param name="flag">An enumeration value.</param>
        /// <returns>
        /// <c>true</c> if the bit field or bit fields that are set in <paramref name="flag"/>
		/// are also set in the current instance; otherwise, <c>false</c>.
        /// </returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool IsSet(this XLCellsUsedOptions self, XLCellsUsedOptions flag)
        {
            return (self & flag) == flag;
        }
    }
}
