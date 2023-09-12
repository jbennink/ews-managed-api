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

using Microsoft.Exchange.WebServices.Data.Misc;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a time zone as used by GetUserAvailabilityRequest.
/// </summary>
internal sealed class LegacyAvailabilityTimeZone : ComplexProperty
{
    private TimeSpan bias;
    private LegacyAvailabilityTimeZoneTime standardTime;
    private LegacyAvailabilityTimeZoneTime daylightTime;

    /// <summary>
    ///     Initializes a new instance of the <see cref="LegacyAvailabilityTimeZone" /> class.
    /// </summary>
    internal LegacyAvailabilityTimeZone()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="LegacyAvailabilityTimeZone" /> class.
    /// </summary>
    /// <param name="timeZoneInfo">The time zone used to initialize this instance.</param>
    internal LegacyAvailabilityTimeZone(TimeZoneInfo timeZoneInfo)
        : this()
    {
        // Availability uses the opposite sign for the bias, e.g. if TimeZoneInfo.BaseUtcOffset = 480 than
        // SerializedTimeZone.Bias must be -480.
        bias = -timeZoneInfo.BaseUtcOffset;

        // To convert TimeZoneInfo into SerializableTimeZone, we need two time changes: one to Standard
        // time, the other to Daylight time. TimeZoneInfo holds a list of adjustment rules that represent
        // the different rules that govern time changes over the years. We need to grab one of those rules
        // to initialize this instance.
        var adjustmentRules = timeZoneInfo.GetAdjustmentRulesEx();

        if (adjustmentRules.Length == 0)
        {
            // If there are no adjustment rules (which is the case for UTC), we have to come up with two
            // dummy time changes which both have a delta of zero and happen at two hard coded dates. This
            // simulates a time zone in which there are no time changes.
            daylightTime = new LegacyAvailabilityTimeZoneTime();
            daylightTime.Delta = TimeSpan.Zero;
            daylightTime.DayOrder = 1;
            daylightTime.DayOfTheWeek = DayOfTheWeek.Sunday;
            daylightTime.Month = 10;
            daylightTime.TimeOfDay = TimeSpan.FromHours(2);
            daylightTime.Year = 0;

            standardTime = new LegacyAvailabilityTimeZoneTime();
            standardTime.Delta = TimeSpan.Zero;
            standardTime.DayOrder = 1;
            standardTime.DayOfTheWeek = DayOfTheWeek.Sunday;
            standardTime.Month = 3;
            standardTime.TimeOfDay = TimeSpan.FromHours(2);
            daylightTime.Year = 0;
        }
        else
        {
            // When there is at least one adjustment rule, we need to grab the last one which is the
            // one that currently applies (TimeZoneInfo stores adjustment rules sorted from oldest to
            // most recent).
            var currentRule = adjustmentRules[adjustmentRules.Length - 1];

            standardTime = new LegacyAvailabilityTimeZoneTime(currentRule.DaylightTransitionEnd, TimeSpan.Zero);

            // Again, TimeZoneInfo and SerializableTime use opposite signs for bias.
            daylightTime = new LegacyAvailabilityTimeZoneTime(
                currentRule.DaylightTransitionStart,
                -currentRule.DaylightDelta
            );
        }
    }

    internal TimeZoneInfo ToTimeZoneInfo()
    {
        if (daylightTime.HasTransitionTime && standardTime.HasTransitionTime)
        {
            var adjustmentRule = AdjustmentRule.CreateAdjustmentRule(
                DateTime.MinValue.Date,
                DateTime.MaxValue.Date,
                -daylightTime.Delta,
                daylightTime.ToTransitionTime(),
                standardTime.ToTransitionTime()
            );

            return TimeZoneExtensions.CreateCustomTimeZone(
                Guid.NewGuid().ToString(),
                -bias,
                "Custom time zone",
                "Standard time",
                "Daylight time",
                new[]
                {
                    adjustmentRule
                }
            );
        }

        // Create no DST time zone
        return TimeZoneExtensions.CreateCustomTimeZone(
            Guid.NewGuid().ToString(),
            -bias,
            "Custom time zone",
            "Standard time"
        );
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Bias:
                bias = TimeSpan.FromMinutes(reader.ReadElementValue<int>());
                return true;
            case XmlElementNames.StandardTime:
                standardTime = new LegacyAvailabilityTimeZoneTime();
                standardTime.LoadFromXml(reader, reader.LocalName);
                return true;
            case XmlElementNames.DaylightTime:
                daylightTime = new LegacyAvailabilityTimeZoneTime();
                daylightTime.LoadFromXml(reader, reader.LocalName);
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    ///     Writes the elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Bias, (int)bias.TotalMinutes);

        standardTime.WriteToXml(writer, XmlElementNames.StandardTime);
        daylightTime.WriteToXml(writer, XmlElementNames.DaylightTime);
    }
}
