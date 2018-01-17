namespace EventLogParser {

    #region Usings
    using Extensions;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Diagnostics;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.Serialization;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using System.Xml.Linq;
    #endregion

    /// <summary>
    /// Base class for properties common to all Windows events
    /// </summary>
    public class EventBase : IValidatableObject {

        #region Members
        #region Static members
        private static readonly string ContainsAlphaCharactersFilter = "[a-zA-Z]*$";
        private static readonly string JunkCharFFFF = new string((char)Int16.Parse("FFFF", NumberStyles.AllowHexSpecifier), count: 1);
        #endregion

        /// <summary>
        /// The event log name
        /// </summary>
        public virtual string Channel { get; protected set; }

        /// <summary>
        /// The DateTime of the event, in Universal Time Coordinated
        /// </summary>
        public virtual DateTime DateTimeUTC { get; protected set; }

        /// <summary>
        /// The event id
        /// </summary>
        public virtual int EventId { get; protected set; }

        /// <summary>
        /// The event record id.  Each event will have a unique record id per machine
        /// </summary>
        public virtual long EventRecordId { get; protected set; }

        /// <summary>
        /// The event source computer
        /// </summary>
        public virtual string EventSourceMachine { get; protected set; }

        /// <summary>
        /// The event level (information, error, warning)
        /// </summary>
        public virtual int Level { get; protected set; }

        #region XElement/Event Data

        protected XElement EventDataElement { get; private set; }

        public IReadOnlyList<XElement> EventDataNameElements { get { return eventDataNameElements; } }
        private List<XElement> eventDataNameElements;

        protected XElement EventSystemElement { get; private set; }

        protected string EventXmlData { get; private set; }

        #endregion
        #endregion

        #region Constructor
        public EventBase() {
            this.Channel = string.Empty;
            this.EventId = -1;
            this.EventRecordId = -1;
            this.EventSourceMachine = string.Empty;
            this.EventXmlData = string.Empty;
            this.Level = -1;
        }

        public EventBase(string eventRecordXml)
            : this() {
            if (string.IsNullOrWhiteSpace(eventRecordXml)) {
                throw new ArgumentNullException("eventRecord");
            }

            this.EventXmlData = eventRecordXml.RemoveControlCharacters().Replace(JunkCharFFFF, " ");

            #region Parse Xml event data
            XElement rootElement = null;
            if (!string.IsNullOrWhiteSpace(this.EventXmlData)) {
                rootElement = XElement.Parse(this.EventXmlData);

                if (rootElement != null) {
                    this.EventSystemElement = rootElement.Elements()
                        .Where(x => x.Name.LocalName == "System")
                        .FirstOrDefault();

                    this.EventDataElement = rootElement.Elements()
                        .Where(x => x.Name.LocalName == "EventData")
                        .FirstOrDefault();

                    if (this.EventDataElement != null) {
                        // <Data Name="SubjectUserSid">S-1-5-18</Data> 
                        this.eventDataNameElements = this.EventDataElement.Elements()
                            .Where(x =>
                                (x != null) && (x.Name != null) && !string.IsNullOrWhiteSpace(x.Name.LocalName)
                                && string.Equals(x.Name.LocalName, "Data", StringComparison.OrdinalIgnoreCase)
                                && x.Attributes()
                            .Any(y =>
                                (y != null) && (y.Name != null) && !string.IsNullOrWhiteSpace(y.Name.LocalName)
                                && string.Equals(y.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase)))
                            .ToList();
                    }

                    if (this.eventDataNameElements == null) {
                        this.eventDataNameElements = new List<XElement>();
                    }
                }
            }

            if (this.EventSystemElement == null) return;
            #endregion

            #region Channel
            this.Channel = this.EventSystemElement.GetElementValue<string>("Channel");
            #endregion

            #region EventSourceMachine
            this.EventSourceMachine = this.EventSystemElement.GetElementValue<string>("Computer");
            if (!string.IsNullOrWhiteSpace(EventSourceMachine)) {
                this.EventSourceMachine = this.EventSourceMachine.Trim();
                if (Regex.IsMatch(this.EventSourceMachine, ContainsAlphaCharactersFilter)) {
                    if (this.EventSourceMachine.IndexOf(".") > 0) {
                        this.EventSourceMachine = this.EventSourceMachine.Substring(0, this.EventSourceMachine.IndexOf("."));
                        if (!string.IsNullOrWhiteSpace(this.EventSourceMachine)) {
                            this.EventSourceMachine = this.EventSourceMachine.Trim().ToUpper();
                        }
                    }
                }
            }
            #endregion

            #region EventId
            this.EventId = this.EventSystemElement.GetElementValue<int>("EventId");
            #endregion

            #region EventRecordId
            this.EventRecordId = this.EventSystemElement.GetElementValue<long>("EventRecordId");
            #endregion

            #region DateTimeUTC
            // the UtcTime in the event data will typically precede the eventRecord.TimeCreated
            // Not using UtcTime due to it may be way off from actual system time          

            var timeCreatedElement = this.EventSystemElement.Elements()
                .Where(x => string.Equals(x.Name.LocalName, "TimeCreated", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();

            if (timeCreatedElement != null) {
                var systemTime = timeCreatedElement.GetAttributeValue<string>("SystemTime");
                if (!string.IsNullOrWhiteSpace(systemTime)) {
                    if (DateTime.TryParse(systemTime.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out DateTime timeCreated)) {
                        this.DateTimeUTC = timeCreated;
                    }
                }
            }
            #endregion

            #region Level
            this.Level = this.EventSystemElement.GetElementValue<int>("Level");
            #endregion

            this.TrimStringValues();
        }
        #endregion

        #region Methods
        /// <summary>
        /// Clears the data structures to conserve memory
        /// </summary>
        public void ClearTemporaryData() {
            this.EventXmlData = null;
            this.EventSystemElement = null;
            this.EventDataElement = null;
            this.eventDataNameElements = null;
        }

        public override string ToString() {
            var info = new StringBuilder();

            info.AppendFormat("{0}; ", this.GetType().Name);
            info.AppendFormat("EventSourceMachine: {0}; ", this.EventSourceMachine ?? "NULL");
            info.AppendFormat("DateTimeUTC: {0}; ", this.DateTimeUTC.YMDHMSFriendly());

            return info.ToString();
        }

        /// <summary>
        /// Performs custom validation of the object
        /// </summary>
        /// <param name="validationContext"></param>
        /// <returns>A list of validation failures</returns>
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext = null) {

            if (this.DateTimeUTC == DateTime.MinValue) {
                yield return new ValidationResult(
                    $"DateTimeUTC is not valid: {this.DateTimeUTC.YMDHMSFriendly()}",
                    new[] { "DateTimeUTC" });
            }

            foreach (var validationResult in this.ValidateProperty(this.EventId, "EventId")) {
                yield return validationResult;
            }

            if (this.EventId < 0) {
                yield return new ValidationResult(
                    $"EventId is not valid: {this.EventId}",
                    new[] { "EventId" });
            }

            if (this.EventRecordId < 1) {
                yield return new ValidationResult(
                    $"EventRecordId is not valid: {this.EventRecordId}",
                    new[] { "EventRecordId" });
            }

            foreach (var validationResult in this.ValidateProperty(this.EventSourceMachine, "EventSourceMachine")) {
                yield return validationResult;
            }

        }

        /// <summary>
        /// Validates the Data Annotations attributes of a property.
        /// </summary>
        /// <param name="propertyToValidate">The property to validate</param>
        /// <param name="memberName">The property name</param>
        /// <returns>The validation results.</returns>
        protected IEnumerable<ValidationResult> ValidateProperty(object propertyToValidate, string memberName) {
            var results = new List<ValidationResult>();

            Validator.TryValidateProperty(
                propertyToValidate, new ValidationContext(this, null, null) { MemberName = memberName }, results);

            return results;
        }

        #endregion
    }
}
