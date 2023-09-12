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

using System.ComponentModel;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the base class for event subscriptions.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class SubscriptionBase
{
    private readonly ExchangeService service;
    private string id;
    private string watermark;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SubscriptionBase" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal SubscriptionBase(ExchangeService service)
    {
        EwsUtilities.ValidateParam(service, "service");

        this.service = service;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SubscriptionBase" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="id">The id.</param>
    internal SubscriptionBase(ExchangeService service, string id)
        : this(service)
    {
        EwsUtilities.ValidateParam(id, "id");

        this.id = id;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SubscriptionBase" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="id">The id.</param>
    /// <param name="watermark">The watermark.</param>
    internal SubscriptionBase(ExchangeService service, string id, string watermark)
        : this(service, id)
    {
        this.watermark = watermark;
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal virtual void LoadFromXml(EwsServiceXmlReader reader)
    {
        id = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.SubscriptionId);

        if (UsesWatermark)
        {
            watermark = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.Watermark);
        }
    }

    /// <summary>
    ///     Gets the session.
    /// </summary>
    /// <value>The session.</value>
    internal ExchangeService Service => service;

    /// <summary>
    ///     Gets the Id of the subscription.
    /// </summary>
    public string Id
    {
        get => id;
        internal set => id = value;
    }

    /// <summary>
    ///     Gets the latest watermark of the subscription. Watermark is always null for streaming subscriptions.
    /// </summary>
    public string Watermark
    {
        get => watermark;
        internal set => watermark = value;
    }

    /// <summary>
    ///     Gets whether or not this subscription uses watermarks.
    /// </summary>
    protected virtual bool UsesWatermark => true;
}
