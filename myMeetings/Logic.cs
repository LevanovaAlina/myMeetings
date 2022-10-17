using System.Globalization;
using Word = Microsoft.Office.Interop.Word;
namespace myMeetings
{
    /// <summary>
    /// Класс, в котором содержится вся логика приложения.
    /// </summary>
     public class Logic : ILogic
    {
        /// <summary>
        /// Поле с названием встречи.
        /// </summary>
        public static string? nameMeeting;

        /// <summary>
        /// Поле с датой встречи.
        /// </summary>
        public static DateTime dateMeeting;

        /// <summary>
        /// Поле с временем начала встречи.
        /// </summary>
        public static TimeSpan startTime;

        /// <summary>
        /// Поле с временем окончания встречи.
        /// </summary>
        public static TimeSpan endTime;

        /// <summary>
        /// Поле с временем оповещения о предстоящей встрече.
        /// </summary>
        public static DateTime? reminderTime = null;

        /// <summary>
        /// Поле с экземпляром класса менеджер встреч.
        /// </summary>
        public static MeetingManager meeting = new MeetingManager();

        /// <summary>
        /// Поле со списком запланированных встреч.
        /// </summary>
        public static List<Meeting> scheduledMeetings = new List<Meeting>();

        /// <summary>
        /// Вывод в консоль меню приложения.
        /// </summary>
        /// <returns>Возвращает значение введенного пользователем числа.</returns>
        public int Menu()
        {
            Console.Clear();
            Console.WriteLine("1 Добавить новую встречу\n2 Изменить встречу\n3 Удалить встречу\n4 Расписание встреч на день" +
                    "\n5 Экспортировать расписание встреч за день в word\n6 Выход");
            return int.Parse(Console.ReadLine());
        }

        /// <summary>
        /// Метод сбора данных о добавляемой встречи через консоль, и добавления ее в расписание.
        /// </summary>
        public void AddMeetings()
        {
            Console.Clear();
            nameMeeting = GetNameByConsole("Введите название встречи:");
            var res = GetDateByConsole("Введите дату встречи в формате дд.мм.гггг (день.месяц.год):");
            while (res == null)
            {
                Console.Clear();
                Console.WriteLine("Дата введена некорректно. Попробуйте снова.");
                Thread.Sleep(5000);
                Console.Clear();
                res = GetDateByConsole("Введите дату встречи в формате дд.мм.гггг (день.месяц.год):");
            }
            dateMeeting = (DateTime)res;
            startTime = GetTimeByConsole("Введите время начала встречи в формате чч:мм (часы:минуты):");
            endTime = GetTimeByConsole("Введите время окончания встречи в формате чч:мм (часы:минуты):");
            Console.WriteLine("Включить уведомление о предстоящей встрече?:");
            Console.WriteLine("1 Да\n2 Нет");
            var k = int.Parse(Console.ReadLine());
            switch (k)
            {
                case 1:
                    reminderTime = GetReminderTimeByConsole("Введите время, за которое вас нужно уведомить, в формате чч:мм(часы:минуты):");
                    break;
                case 2:
                    break;
            }
            Console.Clear();

            if (meeting.TryAddMeeting(nameMeeting, startTime, endTime, dateMeeting, reminderTime))
            {
                Console.WriteLine("Встреча успешно добавлена.");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Время встреч пересекается или встреча запланирована на прошедшую дату. Попробуйте ввести новые данные.");
            }
        }

        /// <summary>
        ///  Метод сбора данных о изменяемой встречи через консоль, и ее изменения в расписании.
        /// </summary>
        public void ChangeMeetings()
        {
            FormingListScheduledMeetingsByDate();
            ListMeetingsByDate(dateMeeting);
            var n = GetSequenceNumberByConsole("Выберите порядковый номер встречи и введите его:");
            ParamMeetingBase(scheduledMeetings, n - 1);
            var l = GetParameterChangeByConsole("Выберите параметр, который хотите изменить:",
                "1 Название\n2 Дата\n3 Время начала встречи\n4 Время окончания встречи\n5 Время оповещения о предстоящей встрече");
            switch (l)
            {
                case 1:
                    nameMeeting = GetNameByConsole("Введите новое название встречи:");
                    break;
                case 2:
                    dateMeeting = GetDateByConsole("Введите новую дату встречи в формате дд.мм.гггг (день.месяц.год):");
                    break;
                case 3:
                    startTime = GetTimeByConsole("Введите новое время начала встречи в формате чч:мм (часы:минуты):");
                    break;
                case 4:
                    endTime = GetTimeByConsole("Введите новое время окончания встречи в формате чч:мм (часы:минуты):");
                    break;
                case 5:
                    reminderTime = GetReminderTimeByConsole("Введите время, за которое вас нужно уведомить, в формате чч:мм (часы:минуты):");
                    break;
            }
            meeting.ChangeMeeting(scheduledMeetings[n - 1].Id, nameMeeting, startTime, endTime, dateMeeting, reminderTime);
            Console.Clear();
            Console.WriteLine("Встреча успешно изменена");
        }

        /// <summary>
        /// Метод удаления встречи из раписания.
        /// </summary>
        public void DeleteMeetings()
        {
            FormingListScheduledMeetingsByDate();
            ListMeetingsByDate(dateMeeting);
            var n = GetSequenceNumberByConsole("Выберите порядковый номер встречи и введите его:");
            meeting.DeleteMeeting(scheduledMeetings[n - 1].Id);
            scheduledMeetings.Remove(scheduledMeetings[n - 1]);
            Console.Clear();
            Console.WriteLine("Встреча успешно удалена");
        }

        /// <summary>
        /// Метод вывода в консоль встреч на заданную дату.
        /// </summary>
        public void MeetingsSchedule()
        {
            FormingListScheduledMeetingsByDate();
            if (scheduledMeetings.Count != 0)
            {
                Console.WriteLine("Расписание на {0}:", dateMeeting.ToShortDateString());
                foreach (Meeting meet in scheduledMeetings)
                {
                    Console.WriteLine("{0} {1} {2} - {3}", meet.Name, meet.DateMeeting.ToShortDateString(), meet.StartTime, meet.EndTime);
                }
            }
            else
            {
                Console.WriteLine("На заданную дату нет запланированных встреч.");
            }
            //Console.WriteLine("Для продолжения нажмите любую клавишу:");
            //Console.ReadKey();
        }

        /// <summary>
        /// Сохранение пути для создания файла со списком встреч и запись этого списка в переменную типа string 
        /// для последующей передачи в метод экспорта в word документ.
        /// </summary>
        /// <param name="path"> Путь, где будет располагаться файл с раписанием встреч.</param>
        public void MeetingsExportWord(string path)
        {
            FormingListScheduledMeetingsByDate();
            var Str = string.Format("Расписание на {0}:", dateMeeting.ToShortDateString());
            foreach (Meeting meet in scheduledMeetings)
            {
                Str += string.Format("\n{0} {1} {2} - {3}", meet.Name, meet.DateMeeting.ToShortDateString(), meet.StartTime, meet.EndTime);
            }
            ExportListMeetingsToWord(Str, path);
            Console.WriteLine("Файл сохранен");
            //Console.WriteLine("Для продолжения нажмите любую клавишу:");
            //Console.ReadKey();
        }

        /// <summary>
        /// Метод получения времени от пользователя в консоли.
        /// </summary>
        /// <param name="str">Строка, которую нужно вывести в консоль.</param>
        /// <returns></returns>
        public TimeSpan GetTimeByConsole(string str)
        {
            Console.Clear();
            Console.WriteLine(str);
            DateTime dt;
            if (!DateTime.TryParseExact(Console.ReadLine(), "HH:mm", CultureInfo.InvariantCulture,
                                                          DateTimeStyles.None, out dt))
            {
                Console.WriteLine("Некорректный ввод данных.");
            }
            TimeSpan time = dt.TimeOfDay;
            Console.Clear();
            //Console.WriteLine("Для продолжения нажмите любую клавишу:");
            //Console.ReadKey();
            return time;
        }

        /// <summary>
        /// Метод получения даты от пользователя в консоли.
        /// </summary>
        /// <param name="str">Строка, которую нужно вывести в консоль.</param>
        /// <returns>Введенная в консоль дата.</returns>
        public dynamic GetDateByConsole(string str)
        {
            Console.WriteLine(str);
            if (DateTime.TryParseExact(Console.ReadLine(), "dd.MM.yyyy", null, DateTimeStyles.None, out var dateMeeting))
            {
                return dateMeeting;
            }
            return null; 
            Console.Clear();
        }

        /// <summary>
        /// Метод получения имени от пользователя в консоли.
        /// </summary>
        /// <param name="str">Строка, которую нужно вывести в консоль.</param>
        /// <returns>Введенная в консоль имя.</returns>
        public string GetNameByConsole(string str)
        {
            Console.WriteLine(str);
            var name = Console.ReadLine();
            Console.Clear();
            return name;
        }

        /// <summary>
        /// Метод получения времени оповещения о встрече от пользователя в консоли.
        /// </summary>
        /// <param name="str">Строка, которую нужно вывести в консоль.</param>
        /// <returns>Введенная в консоль время оповещения о встрече.</returns>
        public DateTime GetReminderTimeByConsole(string str)
        {
            Console.WriteLine(str);
            DateTime dt;
            if (!DateTime.TryParseExact(Console.ReadLine(), "HH:mm", CultureInfo.InvariantCulture,
                                                          DateTimeStyles.None, out dt))
            {
                Console.WriteLine("Некорректный ввод");
                //Console.WriteLine("Для продолжения нажмите любую клавишу:");
                //Console.ReadKey();
            }
            Console.Clear();
            return dt;
        }

        /// <summary>
        /// Метод присваивания новых параметров измененной встречи.
        /// </summary>
        /// <param name="meeting">Имя встречи.</param>
        /// <param name="n">Порядковый номер в списке встреч на заданную дату.</param>
        public void ParamMeetingBase(List<Meeting> meeting, int n)
        {
            nameMeeting = meeting[n].Name;
            dateMeeting = meeting[n].DateMeeting;
            startTime = meeting[n].StartTime;
            endTime = meeting[n].EndTime;
            reminderTime = meeting[n].ReminderTime;
        }

        /// <summary>
        /// Обработка порядкового номера встречи, который был введен пользователем.
        /// </summary>
        /// <param name="str">Строка, которую нужно вывести в консоль.</param>
        /// <param name="strTwo">Строка, которую нужно вывести в консоль.</param>
        /// <returns>Выбранный пользователем порядковый номер.</returns>
        public int GetParameterChangeByConsole(string str, string strTwo)
        {
            Console.WriteLine(str);
            Console.WriteLine(strTwo);
            if (!int.TryParse(Console.ReadLine(), out var l))
            {
                Console.WriteLine("Некорректный ввод");
                //Console.WriteLine("Для продолжения нажмите любую клавишу:");
               // Console.ReadKey();
            }
            Console.Clear();
            return l;
        }

        /// <summary>
        /// Метод для вывода списка встреч на заданную на дату в консоль.
        /// </summary>
        /// <param name="dateMeeting">Дата встречи.</param>
        public void ListMeetingsByDate(DateTime dateMeeting)
        {
            Console.Clear();
            Console.WriteLine("На эту дату {0} назначено встреч: {1}", dateMeeting.ToShortDateString(), scheduledMeetings.Count);
            for (var m = 0; m < scheduledMeetings.Count(); m++)
            {
                Console.WriteLine("{0} {1}", m + 1, scheduledMeetings[m].Name);
            }
        }

        /// <summary>
        /// Метод добавления встреч из общего списка в список встреч на заданную дату.
        /// </summary>
        /// <param name="dateMeeting">Дата встречи.</param>
        /// <returns>Список встреч на заданную дату.</returns>
        public List<Meeting> AddMeetingsInScheduled(DateTime dateMeeting)
        {
            foreach (Meeting meet in meeting.MeetingList)
            {
                if (dateMeeting == meet.DateMeeting)
                {
                    scheduledMeetings.Add(meet);
                }
            }
            Console.Clear();
            return scheduledMeetings;
        }

        /// <summary>
        /// Получение порядкового номера, который был введен пользователем в консоли.
        /// </summary>
        /// <param name="str">Строка, которую нужно вывести в консоль.</param>
        /// <returns>Порядковый номер, который был введен пользователем.</returns>
        public int GetSequenceNumberByConsole(string str)
        {
            Console.WriteLine(str);
            if (!int.TryParse(Console.ReadLine(), out var n))
            {
                Console.WriteLine("Некорректный ввод");
                //Console.WriteLine("Для продолжения нажмите любую клавишу:");
                //Console.ReadKey();
            }
            Console.Clear();
            return n;
        }

        /// <summary>
        /// Метод уведомления о предстоящей встрече.
        /// </summary>
        public void MeetingNotification()
        {
            var i = true;
            while (i)
            {
                var current = new List<Meeting>(meeting.MeetingList).Where(m => m != null);
                foreach (Meeting meet in current)
                {
                    if (meet.ReminderTime != null && !meet.NotificationSent)
                    {
                        var startDateMeeting = meet.DateMeeting.Add(meet.StartTime);
                        var inside = meet.ReminderTime <= DateTime.Now && DateTime.Now <= startDateMeeting;
                        if (inside)
                        {
                            Console.WriteLine("У вас сегодня запланирована встреча: {0} в {1}", meet.Name, meet.StartTime);
                            meet.NotificationSent = true;
                            //Console.WriteLine("Для продолжения нажмите любую клавишу:");
                           // Console.ReadKey();
                            Thread.Sleep(4000);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Метод формирования списка встреч на заданную дату.
        /// </summary>
        public void FormingListScheduledMeetingsByDate()
        {
            dateMeeting = GetDateByConsole("Введите дату встречи в формате дд.мм.гггг (день.месяц.год):");
            scheduledMeetings.Clear();
            scheduledMeetings = AddMeetingsInScheduled(dateMeeting);
        }

        /// <summary>
        /// Метод экспортировани списка встреч в файл word.
        /// </summary>
        /// <param name="str">Расписание встреч в формате string.</param>
        /// <param name="fileName">Путь для создания файла.</param>
        public void ExportListMeetingsToWord(string str, string fileName)
        {
            var wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            wordDoc.Content.Text = str;
            wordApp.Visible = true;
            wordDoc.SaveAs2(fileName);
            wordApp.Application.Documents.Close(fileName);
        }
    }
}



    
