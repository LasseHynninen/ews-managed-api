//-----------------------------------------------------------------------
// <summary>Defines the MailTipsRequested enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the types of requested mail tips.
    /// </summary>
    public enum MailTipsRequested
    {
        /// <summary>
        /// Represents all available mail tips.
        /// </summary>
        All,

        /// <summary>
        /// Represents the Out of Office (OOF) message.
        /// </summary>
        OutOfOfficeMessage,

        /// <summary>
        /// Represents the status for a mailbox that is full.
        /// </summary>
        MailboxFullStatus,

        /// <summary>
        /// Represents a custom mail tip.
        /// </summary>
        CustomMailTip,

        /// <summary>
        /// Represents the count of external members.
        /// </summary>
        ExternalMemberCount,

        /// <summary>
        /// Represents the count of all members.
        /// </summary>
        TotalMemberCount,

        /// <summary>
        /// Represents the maximum message size a recipient can accept.
        /// </summary>
        MaxMessageSize,

        /// <summary>
        /// Indicates whether delivery restrictions will prevent the sender's message from reaching the recipient.
        /// </summary>
        DeliveryRestriction,

        /// <summary>
        /// Indicates whether the sender's message will be reviewed by a moderator.
        /// </summary>
        ModerationStatus,

        /// <summary>
        /// Indicates whether the recipient is invalid.
        /// </summary>
        InvalidRecipient
    }
}