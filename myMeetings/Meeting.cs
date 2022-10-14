namespace myMeetings
{
    public class Meeting
    {
        public readonly string Id;
        public string Name;
        public TimeSpan StartTime;
        public TimeSpan EndTime;
        public DateTime DateMeeting;
        public DateTime? ReminderTime;
        public bool NotificationSent = false;
        //public DateTime DateAndTimeReminder
        public Meeting (string name, TimeSpan startTime, TimeSpan endTime, DateTime dateMeeting, DateTime? reminderTime)
        {
            this.Id = Guid.NewGuid().ToString();
            this.Name = name;
            this.StartTime = startTime;
            this.EndTime = endTime;
            if(!(reminderTime == null))
            {   
                var reminderTimeNotNull = (DateTime)reminderTime;
                this.ReminderTime = new DateTime(dateMeeting.Year, dateMeeting.Month, dateMeeting.Day);
                var s = startTime.Add(-reminderTimeNotNull.TimeOfDay);
                this.ReminderTime += s;
            }
            this.DateMeeting = dateMeeting;
        }
    }
}
