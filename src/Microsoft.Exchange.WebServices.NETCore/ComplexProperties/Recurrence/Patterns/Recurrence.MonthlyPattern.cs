/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data;

/// <content>
///     Contains nested type Recurrence.MonthlyPattern.
/// </content>
public abstract partial class Recurrence
{
    /// <summary>
    ///     Represents a recurrence pattern where each occurrence happens on a specific day a specific number of
    ///     months after the previous one.
    /// </summary>
    public sealed class MonthlyPattern : IntervalPattern
    {
        private int? dayOfMonth;

        /// <summary>
        ///     Initializes a new instance of the <see cref="MonthlyPattern" /> class.
        /// </summary>
        public MonthlyPattern()
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="MonthlyPattern" /> class.
        /// </summary>
        /// <param name="startDate">The date and time when the recurrence starts.</param>
        /// <param name="interval">The number of months between each occurrence.</param>
        /// <param name="dayOfMonth">The day of the month when each occurrence happens.</param>
        public MonthlyPattern(DateTime startDate, int interval, int dayOfMonth)
            : base(startDate, interval)
        {
            DayOfMonth = dayOfMonth;
        }

        /// <summary>
        ///     Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal override string XmlElementName => XmlElementNames.AbsoluteMonthlyRecurrence;

        /// <summary>
        ///     Write properties to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void InternalWritePropertiesToXml(EwsServiceXmlWriter writer)
        {
            base.InternalWritePropertiesToXml(writer);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DayOfMonth, DayOfMonth);
        }

        /// <summary>
        ///     Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (base.TryReadElementFromXml(reader))
            {
                return true;
            }

            switch (reader.LocalName)
            {
                case XmlElementNames.DayOfMonth:
                    dayOfMonth = reader.ReadElementValue<int>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        ///     Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();

            if (!dayOfMonth.HasValue)
            {
                throw new ServiceValidationException(Strings.DayOfMonthMustBeBetween1And31);
            }
        }

        /// <summary>
        ///     Gets or sets the day of the month when each occurrence happens. DayOfMonth must be between 1 and 31.
        /// </summary>
        public int DayOfMonth
        {
            get => GetFieldValueOrThrowIfNull(dayOfMonth, "DayOfMonth");

            set
            {
                if (value < 1 || value > 31)
                {
                    throw new ArgumentOutOfRangeException("DayOfMonth", Strings.DayOfMonthMustBeBetween1And31);
                }

                SetFieldValue(ref dayOfMonth, value);
            }
        }

        /// <summary>
        ///     Checks if two recurrence objects are identical.
        /// </summary>
        /// <param name="otherRecurrence">The recurrence to compare this one to.</param>
        /// <returns>true if the two recurrences are identical, false otherwise.</returns>
        public override bool IsSame(Recurrence otherRecurrence)
        {
            return base.IsSame(otherRecurrence) && dayOfMonth == ((MonthlyPattern)otherRecurrence).dayOfMonth;
        }
    }
}
