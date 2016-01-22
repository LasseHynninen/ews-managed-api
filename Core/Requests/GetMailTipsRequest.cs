﻿namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a GetMailTips request.
    /// </summary>
    internal sealed class GetMailTipsRequest : SimpleServiceRequestBase
    {
        //https://msdn.microsoft.com/en-us/library/office/dd877060(v=exchg.140).aspx [GetMailTips Operation][2010]
        //https://msdn.microsoft.com/en-us/library/office/dd877060(v=exchg.150).aspx [GetMailTips Operation][2013]
        private MailTipsRequested requestedMailTips;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetMailTipsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetMailTipsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name</returns>
        internal override string GetXmlElementName() { return XmlElementNames.GetMailTips; }

        /// <summary>Writes XML elements.</summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            SendingAs.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.SendingAs);

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Recipients);
            foreach (var mbox in Recipients)
            {
                mbox.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.Mailbox);
            }
            writer.WriteEndElement(); // </Recipients>

            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MailTipsRequested, requestedMailTips);
        }

        /// <summary>Gets the name of the response XML element.</summary>
        /// <returns>XML element name</returns>
        internal override string GetResponseXmlElementName() { return XmlElementNames.GetMailTipsResponse; }

        /// <summary>Parses the response.</summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetMailTipsResults serviceResponse = new GetMailTipsResults();
            serviceResponse.LoadFromXml(reader, XmlElementNames.GetMailTipsResponse);
            //if (serviceResponse.ErrorCode != ServiceError.NoError)
            return serviceResponse;
        }

        /// <summary>Gets the request version.</summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010;
        }

        /// <summary>Executes this request.</summary>
        /// <returns>Service response.</returns>
        internal GetMailTipsResults Execute()
        {
            return (GetMailTipsResults)this.InternalExecute();
        }

        /// <summary>Gets or sets the attendees.</summary>
        public EmailAddress SendingAs { get; set; }

        /// <summary>Gets or sets the requested MailTips.</summary>
        public MailTipsRequested MailTipsRequested
        {
            get { return this.requestedMailTips; }
            set { this.requestedMailTips = value; }
        }

        /// <summary>
        /// Gets or sets who are the recipients/targets whose MailTips we are interested in.
        /// </summary>
        public Mailbox[] Recipients { get; set; }
    }
}