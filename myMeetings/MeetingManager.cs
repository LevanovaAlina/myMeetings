namespace myMeetings
{
    public class MeetingManager
    {
        public List<Meeting> MeetingList = new List<Meeting>();
        public bool TryAddMeeting(string name, TimeSpan startTime, TimeSpan endTime, DateTime dateMeeting, DateTime? reminderTime)
        {
            var canAdd = true;
            foreach(Meeting meeting in MeetingList)
            {
                var crossed = meeting.DateMeeting == dateMeeting &&
                    ((startTime >= meeting.StartTime && startTime <= meeting.EndTime) ||
                    (endTime >= meeting.StartTime && endTime <= meeting.EndTime));
                if (crossed)
                {
                    canAdd = false;
                    break;
                }
            }
            if(canAdd)
            {
                MeetingList.Add(new Meeting(name, startTime, endTime, dateMeeting, reminderTime));
                return true;
            }
            return false;   
        }

        public void ChangeMeeting(string id, string name, TimeSpan startTime, TimeSpan endTime, DateTime dateMeeting, DateTime? reminderTime)
        {
            foreach(Meeting meeting in MeetingList)
            {
                if (meeting.Id == id)
                {
                    meeting.Name = name;
                    meeting.StartTime = startTime;
                    meeting.EndTime = endTime;
                    meeting.DateMeeting = dateMeeting;
                    meeting.ReminderTime = reminderTime;
                }
            }
        }

        public void DeleteMeeting(string id)
        {
            for (var i = MeetingList.Count - 1; i >= 0; i--) 
            {
                if (MeetingList[i].Id == id)
                {
                    MeetingList.Remove(MeetingList[i]);
                }
            }
        }
    }
}
