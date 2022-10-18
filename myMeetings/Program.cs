namespace myMeetings
{
    //TODO: Приложение можно дополнить в будущем: ui, создать интерактивное меню,
    //всплывающее окно-уведомление о встрече

    /// <summary>
    /// Основной класс программы.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            var logic = new Logic();
            var ended = false;
            var thread = new Thread(logic.MeetingNotification);
            thread.Start();
            while (!ended)
            {
                var i = logic.Menu();
                switch (i)
                {
                    case 1:
                       logic.AddMeetings();
                        Console.ReadKey();
                        break;
                    case 2:
                        logic.ChangeMeetings();
                        Console.ReadKey();
                        break;
                    case 3:
                        logic.DeleteMeetings();
                        Console.ReadKey();
                        break;
                    case 4:
                        logic.MeetingsSchedule();
                        Console.ReadKey();
                        break;
                    case 5:
                        Console.WriteLine("Введите путь для сохранения файла:");
                        var path = Console.ReadLine();
                        logic.MeetingsExportWord(path);
                        break;
                    case 6:
                        Console.Clear();
                        Console.WriteLine("Выход осущетсвлен.");
                        ended = true;
                        break;
                    default:
                        Console.Clear();
                        Console.WriteLine("Ошибка работы приложения.");
                        break;
                }
            }
        }
    }
}