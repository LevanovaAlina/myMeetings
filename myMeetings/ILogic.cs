
namespace myMeetings
{
    /// <summary>
    /// Интерфейс логики приложения.
    /// </summary>
    public interface ILogic
    {
        public static string? nameMeeting;
        public static DateTime? dateMeeting;
        public static TimeSpan startTime;
        public static TimeSpan endTime;
        public static DateTime? reminderTime;
        public static MeetingManager meeting;
        public static List<Meeting> scheduledMeetings;
        public abstract int Menu();
        public abstract void AddMeetings();
        public abstract void ChangeMeetings();
        public abstract void DeleteMeetings();
        public abstract void MeetingsSchedule();
        public abstract void MeetingsExportWord(string path);
        public abstract TimeSpan GetTimeByConsole(string str);
        public abstract dynamic GetDateByConsole(string str);
        public abstract string GetNameByConsole(string str);
        public abstract DateTime GetReminderTimeByConsole(string str);
        public abstract void ParamMeetingBase(List<Meeting> meeting, int n);
        public abstract int GetParameterChangeByConsole(string str, string strTwo);
        public abstract void ListMeetingsByDate(DateTime dateMeeting);
        public abstract List<Meeting> AddMeetingsInScheduled(DateTime dateMeeting);
        public abstract int GetSequenceNumberByConsole(string str);
        public abstract void MeetingNotification();
        public abstract void FormingListScheduledMeetingsByDate();
        public abstract void ExportListMeetingsToWord(string str, string fileName);
    }
}
