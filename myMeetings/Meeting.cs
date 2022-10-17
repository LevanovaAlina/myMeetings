namespace myMeetings
{
    /// <summary>
    /// Класс сущности: встреча.
    /// </summary>
    public class Meeting
    {
        /// <summary>
        /// Поле с id.
        /// </summary>
        public readonly Guid Id;

        /// <summary>
        /// Имя встречи.
        /// </summary>
        public string Name;

        /// <summary>
        /// Время начала встречи.
        /// </summary>
        public TimeSpan StartTime;

        /// <summary>
        /// Время окончания встречи.
        /// </summary>
        public TimeSpan EndTime;

        /// <summary>
        /// Дата встречи.
        /// </summary>
        public DateTime DateMeeting;

        /// <summary>
        /// Время, за которое нужно уведомить пользователя о предстоящей встрече.
        /// </summary>
        private DateTime? reminderTime;

        /// <summary>
        /// Было ли отправлено уведомление о встрече.
        /// </summary>
        public bool NotificationSent = false;

        /// <summary>
        /// Время, в которое придет уведомление о встрече.
        /// </summary>
        public DateTime? ReminderTime
        {
            get
            {
                if (!(this.reminderTime == null))
                {
                    var reminderTimeNotNull = (DateTime)this.reminderTime;
                    var reminderTime = new DateTime(this.DateMeeting.Year, this.DateMeeting.Month, this.DateMeeting.Day);
                    var s = this.StartTime.Add(-reminderTimeNotNull.TimeOfDay);
                    reminderTime += s;
                    return reminderTime;
                }
                return null;    
            }
            set 
            { 
                this.reminderTime = value; 
            }
        }

        /// <summary>
        /// Заполнение полей с данными о встрече.
        /// </summary>
        /// <param name="name">Название встречи.</param>
        /// <param name="startTime">Время начала встречи.</param>
        /// <param name="endTime">Время окончания встречи.</param>
        /// <param name="dateMeeting">Дата встречи.</param>
        /// <param name="reminderTime">Время, за которое нужно уведомить о встрече.</param>
        public Meeting (string name, TimeSpan startTime, TimeSpan endTime, DateTime dateMeeting, DateTime? reminderTime)
        {
            this.Id = Guid.NewGuid();
            this.Name = name;
            this.StartTime = startTime;
            this.EndTime = endTime;
            this.ReminderTime = reminderTime;
            this.DateMeeting = dateMeeting;
        }
    }
}
