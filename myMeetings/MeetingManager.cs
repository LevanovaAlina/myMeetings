namespace myMeetings
{
    /// <summary>
    /// Класс для работы со списком встреч.
    /// </summary>
    public class MeetingManager
    {
        /// <summary>
        /// Список встреч.
        /// </summary>
        public readonly List<Meeting> MeetingList = new List<Meeting>();
        
        /// <summary>
        /// Метод добавления встречи в список.
        /// </summary>
        /// <param name="name">Название встречи.</param>
        /// <param name="startTime">Время начала встречи.</param>
        /// <param name="endTime">Время окончания встречи.</param>
        /// <param name="dateMeeting">Дата встречи.</param>
        /// <param name="reminderTime">Время, за которое нужно уведомить о встрече.</param>
        /// <returns>Успешность добавления встречи.</returns>
        public bool TryAddMeeting(string name, TimeSpan startTime, TimeSpan endTime, DateTime dateMeeting, DateTime? reminderTime)
        {
            if (dateMeeting.Date < DateTime.Now.Date)
            {
                return false;
            }    
            foreach (Meeting meeting in MeetingList)
            {
                var crossed = meeting.DateMeeting == dateMeeting &&
                    ((startTime >= meeting.StartTime && startTime <= meeting.EndTime) ||
                    (endTime >= meeting.StartTime && endTime <= meeting.EndTime));
                if (crossed)
                {
                    return false;
                }
            }
            MeetingList.Add(new Meeting(name, startTime, endTime, dateMeeting, reminderTime));
            return true;
        }

        /// <summary>
        /// Метод изменения встречи.
        /// </summary>
        /// <param name="id">Id встречи.</param>
        /// <param name="name">Название встречи.</param>
        /// <param name="startTime">Время начала встречи.</param>
        /// <param name="endTime">Время окончания встречи.</param>
        /// <param name="dateMeeting">Дата встречи.</param>
        /// <param name="reminderTime">Время, за которое нужно уведомить о встрече.</param>
        /// <returns>Успешность добавления встречи.</param>
        public void ChangeMeeting(Guid id, string name, TimeSpan startTime, TimeSpan endTime, DateTime dateMeeting, DateTime? reminderTime)
        {
            foreach (Meeting meeting in MeetingList)
            {
                if (meeting.Id == id)
                {
                    meeting.Name = name;
                    meeting.StartTime = startTime;
                    meeting.EndTime = endTime;
                    meeting.DateMeeting = dateMeeting;
                    meeting.ReminderTime = reminderTime;
                    break;
                }
            }
        }

        /// <summary>
        /// Метод удаления встречи.
        /// </summary>
        /// <param name="id">Id встречи.</param>
        public void DeleteMeeting(Guid id)
        {
            for (var i = MeetingList.Count - 1; i >= 0; i--)
            {
                if (MeetingList[i].Id == id)
                {
                    MeetingList.Remove(MeetingList[i]);
                    break;
                }
            }
        }
    }
}
