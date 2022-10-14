using System.Globalization;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
namespace myMeetings
{
    //TODO: Приложение можно дополнить в будущем: ui, создать интерактивное меню,
    //всплывающее окно-уведомление о встрече, добавить обработку ошибок при некорректном вводе от пользователя,
    //поправить многопоточность,
    //добавить обработку ошибок при пересечении времени встреч,
    //сделать рефакторинг кода,
    //отделить модель от логики,
    //выделить часть кода из main в отдельные классы,
    //написать модульные тесты
    class Program
    {
        public static string? nameMeeting;
        public static DateTime dateMeeting;
        public static TimeSpan startTime;
        public static TimeSpan endTime;
        public static DateTime? reminderTime = null;
        public static MeetingManager meeting = new MeetingManager();
        public static List<Meeting> scheduledMeetings = new List<Meeting>();
        public static object locker = new();
        public static AutoResetEvent waitHandler = new AutoResetEvent(true);
        public static string path = "D:\\word\\meetingsList.docx"; //путь для будущего файла для экспорта
        static void Main(string[] args)
        {
            var ended = false;
            var thread = new Thread(MeetingNotification);
            thread.Priority = ThreadPriority.BelowNormal;
            thread.Start();
            while (!ended)
            {
                Console.Clear();
                Console.WriteLine("1 Добавить новую встречу\n2 Изменить встречу\n3 Удалить встречу\n4 Расписание встреч на день" +
                    "\n5 Экспортировать расписание встреч за день в word\n6 Выход");
                int i = int.Parse(Console.ReadLine());
                switch (i)
                {
                    case 1:
                        Console.Clear();
                        nameMeeting = GetNameByConsole("Введите название встречи:");
                        dateMeeting = GetDateByConsole("Введите дату встречи в формате дд.мм.гггг (день.месяц.год):");
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
                            Console.WriteLine("Встреча успешно добавлена");
                            Console.ReadKey();
                        }
                        else
                        {
                            Console.WriteLine("Ошибка добавления встречи");
                        }
                        break;
                    case 2:
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
                        break;
                    case 3:
                        FormingListScheduledMeetingsByDate();
                        n = GetSequenceNumberByConsole("Выберите порядковый номер встречи и введите его:");
                        meeting.DeleteMeeting(scheduledMeetings[n - 1].Id);
                        scheduledMeetings.Remove(scheduledMeetings[n - 1]);
                        Console.Clear();
                        Console.WriteLine("Встреча успешно удалена");
                        break;
                    case 4:
                        FormingListScheduledMeetingsByDate();
                        Console.WriteLine("Расписание на {0}:", dateMeeting.ToShortDateString());
                        foreach (Meeting meet in scheduledMeetings)
                        {
                            Console.WriteLine("{0} {1} {2} - {3} {4}", meet.Name, meet.DateMeeting.ToShortDateString(), meet.StartTime, meet.EndTime, meet.ReminderTime);
                        }
                        Thread.Sleep(10000);
                        break;
                    case 5:
                        FormingListScheduledMeetingsByDate();
                        var Str = string.Format("Расписание на {0}:", dateMeeting.ToShortDateString());
                        foreach (Meeting meet in scheduledMeetings)
                        {
                            Str += string.Format("\n{0} {1} {2} - {3} {4}", meet.Name, meet.DateMeeting.ToShortDateString(), meet.StartTime, meet.EndTime, meet.ReminderTime);
                        }                      
                        ExportListMeetingsToWord(Str, path);
                        Console.WriteLine("Файл сохранен");
                        Console.ReadKey();
                        break;


                    case 6:
                        Console.WriteLine("Выход");
                        ended = true;
                        break;
                    default:
                        Console.WriteLine("Ошибка");
                        break;
                }
            }
            Thread.Sleep(400);
        }

        private static TimeSpan GetTimeByConsole(string str)
        {
            Console.WriteLine(str);
            DateTime dt;
            if (!DateTime.TryParseExact(Console.ReadLine(), "HH:mm", CultureInfo.InvariantCulture,
                                                          DateTimeStyles.None, out dt))
            {
                Console.WriteLine("Неккоректный ввод");
            }
            TimeSpan time = dt.TimeOfDay;
            Console.Clear();
            return time;
        }

        private static DateTime GetDateByConsole(string str)
        {
            Console.WriteLine(str);
            DateTime.TryParseExact(Console.ReadLine(), "dd.MM.yyyy", null, DateTimeStyles.None, out var dateMeeting);
            Console.Clear();
            return dateMeeting;
        }

        private static string GetNameByConsole(string str)
        {
            Console.WriteLine(str);
            var name = Console.ReadLine();
            Console.Clear();
            return name;
        }

        private static DateTime GetReminderTimeByConsole(string str)
        {
            Console.WriteLine(str);
            DateTime dt;
            if (!DateTime.TryParseExact(Console.ReadLine(), "HH:mm", CultureInfo.InvariantCulture,
                                                          DateTimeStyles.None, out dt))
            {
                Console.WriteLine("Неккоректный ввод");
            }
            Console.Clear();
            return dt;
        }

        private static void ParamMeetingBase(List<Meeting> meeting, int n)
        {
            nameMeeting = meeting[n].Name;
            dateMeeting = meeting[n].DateMeeting;
            startTime = meeting[n].StartTime;
            endTime = meeting[n].EndTime;
            reminderTime = meeting[n].ReminderTime;
        }
        private static int GetParameterChangeByConsole(string str, string strTwo)
        {
            Console.WriteLine(str);
            Console.WriteLine(strTwo);
            if (!int.TryParse(Console.ReadLine(), out var l))
            {
                Console.WriteLine("Неккоректный ввод");
            }
            Console.Clear();
            return l;
        }

        private static void ListMeetingsByDate(DateTime dateMeeting)
        {
            Console.Clear();
            Console.WriteLine("На эту дату {0} назначено встреч: {1}", dateMeeting.ToShortDateString(), scheduledMeetings.Count);
            for (var m = 0; m < scheduledMeetings.Count(); m++)
            {
                Console.WriteLine("{0} {1}", m + 1, scheduledMeetings[m].Name);
            }
        }
        private static List<Meeting> AddMeetingsInScheduled(DateTime dateMeeting)
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

        private static int GetSequenceNumberByConsole(string str)
        {
            Console.WriteLine(str);
            if (!int.TryParse(Console.ReadLine(), out var n))
            {
                Console.WriteLine("Неккоректный ввод");
            }
            Console.Clear();
            return n;
        }

        private static void MeetingNotification()
        {
            var i = true;
            while (i)
            {
                waitHandler.WaitOne();
                foreach (Meeting meet in meeting.MeetingList.Where(m => m != null && m.NotificationSent == false))
                {
                    if (meet.ReminderTime != null)
                    {
                        var startDateMeeting = meet.DateMeeting.Add(meet.StartTime);
                        var inside = meet.ReminderTime <= DateTime.Now && DateTime.Now <= startDateMeeting;
                        if (inside)
                        {
                            Console.WriteLine("У вас сегодня запланирована встреча: {0} в {1}", meet.Name, meet.StartTime);
                            meet.NotificationSent = true;
                            Thread.Sleep(4000);
                        }
                    }
                }
                waitHandler.Set();
            }
        }

        private static void FormingListScheduledMeetingsByDate()
        {
            dateMeeting = GetDateByConsole("Введите дату встречи в формате дд.мм.гггг (день.месяц.год):");
            scheduledMeetings.Clear();
            scheduledMeetings = AddMeetingsInScheduled(dateMeeting);
        }

         private static void ExportListMeetingsToWord(string str, string fileName)
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
