﻿// ---------------------------------------------------------------------------
// <copyright file="ComputedInsightValueProperty.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements the class for computed insight value property.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a computed insight value.
/// </summary>
public sealed class ComputedInsightValueProperty : ComplexProperty
{
    private string key;
    private string value;

    /// <summary>
    ///     Gets or sets the Key
    /// </summary>
    public string Key
    {
        get => key;

        set => SetFieldValue(ref key, value);
    }

    /// <summary>
    ///     Gets or sets the Value
    /// </summary>
    public string Value
    {
        get => value;

        set => SetFieldValue(ref this.value, value);
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">XML reader</param>
    /// <returns>Whether the element was read</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Key:
                Key = reader.ReadElementValue();
                break;
            case XmlElementNames.Value:
                Value = reader.ReadElementValue();
                break;
            default:
                return false;
        }

        return true;
    }
}
