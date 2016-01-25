using System;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Out of Office part from the MailTips response
    /// </summary>
    public sealed class MailTipOutOfOffice : ServiceResponse
    {
        private string message;
        private DateTime? startTime;
        private DateTime? endTime;

        internal MailTipOutOfOffice()
            : base()
        {
        }

        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.OutOfOffice);
            reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ReplyBody);
            message = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Message);
            reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.ReplyBody);
            reader.Read();
            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Duration))
            {
                reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Duration);
                var start = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.StartTime);
                startTime = DateTime.Parse(start);
                var end = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.EndTime);
                endTime = DateTime.Parse(end);
                reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.Duration);
            }
            reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.OutOfOffice);
            reader.Read();
        }

        /// <summary>
        /// Gets the message contained in Out Of Office-message
        /// </summary>
        public string Message { get { return message; } }

        /// <summary>
        /// Gets the start time contained in Out Of Office-message if it exists.
        /// </summary>
        public DateTime? StartTime { get { return startTime;} }

        /// <summary>
        /// Gets the end time contained in Out Of Office-message if it exists.
        /// </summary>
        public DateTime? EndTime { get { return endTime; } }
    }
}