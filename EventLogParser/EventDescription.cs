namespace EventLogParser {

    #region Usings
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    #endregion

    [Serializable]
    public class EventDescription {

        #region Members
        /// <summary>
        /// The event description
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// The event id
        /// </summary>
        public int EventId { get; set; }

        /// <summary>
        /// The event log name
        /// </summary>
        public string EventLog { get; set; }
        #endregion

        #region Methods
        public bool IsValid() {
            return !string.IsNullOrWhiteSpace(this.EventLog)
                && (this.EventId >= 0);
        }

        public override string ToString() {
            var info = new StringBuilder();

            info.Append($"{nameof(this.EventLog)}: {this.EventLog ?? "NULL"}; ");
            info.Append($"{nameof(this.EventId)}: {this.EventId}; ");
            info.Append($"{nameof(this.Description)}: {this.Description ?? "NULL"}; ");

            return info.ToString();
        }
        #endregion
    }
}
