using PowerShellPowered.ExtensionMethods;

namespace PowerShellPowered.PowerShellConnect.Entities
{
    public class ShellDataStreams
    {
        public ShellDebugData Debug { get; set; }
        public ShellErrorData Error { get; set; }
        public string GenericErrorMessage { get; set; }
        public ShellWarningData Warning { get; set; }
        public ShellVerboseData Verbose { get; set; }
        public ShellProgressData Progress { get; set; }
        public ShellStreamDataType StreamDataType { get; set; }

        public static ShellDataStreams CreateProgressDataStream(int activityId, string activity, string statusDescription, int percentComplete)
        {
            ShellDataStreams sds = new ShellDataStreams();
            sds.StreamDataType = ShellStreamDataType.Progress;
            sds.Progress = new ShellProgressData() { ActivityId = activityId, Activity = activity, StatusDescription = statusDescription };
            return sds;
        }

        public static ShellDataStreams CreateGenericError(string genericErrors)
        {
            return new ShellDataStreams() { GenericErrorMessage = genericErrors, StreamDataType = ShellStreamDataType.GenericError };
        }


    }

    public abstract class ShellStreamData
    {
        /// <summary>
        /// ToString() _value of record type
        /// </summary>
        public string DetailedData { get; set; }
        public string Message { get; set; }
        public bool ExpectingInput { get; set; }
        public string Line { get; set; }
        public string PositionMessage { get; set; }
        public int ScriptLineNumber { get; set; }

        public override string ToString()
        {
            return DetailedData.IsNullOrEmpty() ? Message : DetailedData;
        }
    }

    public class ShellDebugData : ShellStreamData { }

    public class ShellVerboseData : ShellStreamData { }

    public class ShellWarningData : ShellStreamData
    {
        public string FullyQualifiedWarningId { get; set; }
    }

    public class ShellErrorData : ShellStreamData
    {
        public ShellErrorData() { }
        public ShellErrorData(string message) { Message = message; }

        public string FullyQualifiedErrorId { get; set; }
        public string Category { get; set; }
        public string ExceptionMessage { get; set; }
    }

    public class ShellProgressData
    {
        /// <summary>
        /// ToString() _value of verbose data
        /// </summary>
        public string ToStringValue { get; set; }

        public string Activity { get; set; }
        public int ActivityId { get; set; }
        public string CurrentOperation { get; set; }
        public int ParentActivityId { get; set; }
        public int PercentComplete { get; set; }
        public bool IsCompleted { get; set; }
        //public PSProgressDataType RecordType { get; set; }
        public int SecondsRemaining { get; set; }
        public string StatusDescription { get; set; }

        public override string ToString()
        {
            return ToStringValue;
        }
    }
}
