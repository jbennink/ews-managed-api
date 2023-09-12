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
///     Represents an EmailUserEntity object.
/// </summary>
public sealed class EmailUserEntity : ComplexProperty
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailUserEntity" /> class.
    /// </summary>
    internal EmailUserEntity()
    {
        Namespace = XmlNamespace.Types;
    }

    /// <summary>
    ///     Gets the EmailUser entity Name.
    /// </summary>
    public string Name { get; internal set; }

    /// <summary>
    ///     Gets the EmailUser entity UserId.
    /// </summary>
    public string UserId { get; internal set; }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.NlgName:
                Name = reader.ReadElementValue();
                return true;

            case XmlElementNames.NlgUserId:
                UserId = reader.ReadElementValue();
                return true;

            default:
                return base.TryReadElementFromXml(reader);
        }
    }
}
