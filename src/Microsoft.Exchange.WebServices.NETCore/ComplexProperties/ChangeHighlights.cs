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

/// <summary>
///     Encapsulates information on the changehighlights of a meeting request.
/// </summary>
public sealed class ChangeHighlights : ComplexProperty
{
    private bool hasLocationChanged;
    private string location;
    private bool hasStartTimeChanged;
    private DateTime start;
    private bool hasEndTimeChanged;
    private DateTime end;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ChangeHighlights" /> class.
    /// </summary>
    internal ChangeHighlights()
    {
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
            case XmlElementNames.HasLocationChanged:
                hasLocationChanged = reader.ReadElementValue<bool>();
                return true;
            case XmlElementNames.Location:
                location = reader.ReadElementValue();
                return true;
            case XmlElementNames.HasStartTimeChanged:
                hasStartTimeChanged = reader.ReadElementValue<bool>();
                return true;
            case XmlElementNames.Start:
                start = reader.ReadElementValueAsDateTime().Value;
                return true;
            case XmlElementNames.HasEndTimeChanged:
                hasEndTimeChanged = reader.ReadElementValue<bool>();
                return true;
            case XmlElementNames.End:
                end = reader.ReadElementValueAsDateTime().Value;
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    ///     Gets a value indicating whether the location has changed.
    /// </summary>
    public bool HasLocationChanged => hasLocationChanged;

    /// <summary>
    ///     Gets the old location
    /// </summary>
    public string Location => location;

    /// <summary>
    ///     Gets a value indicating whether the the start time has changed.
    /// </summary>
    public bool HasStartTimeChanged => hasStartTimeChanged;

    /// <summary>
    ///     Gets the old start date and time of the meeting.
    /// </summary>
    public DateTime Start => start;

    /// <summary>
    ///     Gets a value indicating whether the the end time has changed.
    /// </summary>
    public bool HasEndTimeChanged => hasEndTimeChanged;

    /// <summary>
    ///     Gets the old end date and time of the meeting.
    /// </summary>
    public DateTime End => end;
}
