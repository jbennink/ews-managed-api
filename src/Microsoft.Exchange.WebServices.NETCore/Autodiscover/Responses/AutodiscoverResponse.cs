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

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Represents the base class for all responses returned by the Autodiscover service.
/// </summary>
public abstract class AutodiscoverResponse
{
    private AutodiscoverErrorCode errorCode;
    private string errorMessage;
    private Uri redirectionUrl;

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverResponse" /> class.
    /// </summary>
    internal AutodiscoverResponse()
    {
        errorCode = AutodiscoverErrorCode.NoError;
        errorMessage = Strings.NoError;
    }

    /// <summary>
    ///     Loads response from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="endElementName">End element name.</param>
    internal virtual void LoadFromXml(EwsXmlReader reader, string endElementName)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.ErrorCode:
                ErrorCode = reader.ReadElementValue<AutodiscoverErrorCode>();
                break;
            case XmlElementNames.ErrorMessage:
                ErrorMessage = reader.ReadElementValue();
                break;
        }
    }


    #region Properties

    /// <summary>
    ///     Gets the error code that was returned by the service.
    /// </summary>
    public AutodiscoverErrorCode ErrorCode
    {
        get => errorCode;
        internal set => errorCode = value;
    }

    /// <summary>
    ///     Gets the error message that was returned by the service.
    /// </summary>
    /// <value>The error message.</value>
    public string ErrorMessage
    {
        get => errorMessage;
        internal set => errorMessage = value;
    }

    /// <summary>
    ///     Gets or sets the redirection URL.
    /// </summary>
    /// <value>The redirection URL.</value>
    internal Uri RedirectionUrl
    {
        get => redirectionUrl;
        set => redirectionUrl = value;
    }

    #endregion
}
