using Microsoft.Win32.TaskScheduler;
using Task = Microsoft.Win32.TaskScheduler.Task;

namespace Ebankp.ScheduleTask
{
    public class Program
    {
        /// <summary>
        /// 工作排程CURD製成範例
        /// </summary>
        /// <author>BuzzLightYear</author>
        /// <param name="args"></param>
        /// <returns></returns>
        public static async System.Threading.Tasks.Task Main(string[] args)
        {
            //CURD第一參數
            string jobCode = args[0];
            //工作名稱
            string jobName = args[1];

            TaskService ts = TaskService.Instance;
            switch (jobCode)
            {
                //新增工作排程
                case "AddNewTask_Notepad++":

                    //名稱
                    //string taskName = "Open Notepad++";
                    //新增
                    TaskDefinition td = TaskService.Instance.NewTask();
                    //作者
                    td.RegistrationInfo.Author = "Buzz LightYear";
                    //描述
                    td.RegistrationInfo.Description = "Sample task opening notepadd++";

                    //觸發程序
                    //翻譯:每兩天當下時間的每小時重複一次持續四小時
                    DailyTrigger dt = (DailyTrigger)td.Triggers.Add(new DailyTrigger(2));
                    dt.Repetition.Duration = TimeSpan.FromHours(4);
                    dt.Repetition.Interval = TimeSpan.FromHours(1);

                    //觸發程序
                    //翻譯:從今天開始,每個星期的星期五上午兩點
                    td.Triggers.Add(new WeeklyTrigger
                    {
                        StartBoundary = DateTime.Today
                      + TimeSpan.FromHours(2),
                        DaysOfWeek = DaysOfTheWeek.Friday
                    });

                    //動作
                    td.Actions.Add(new ExecAction(@"C:\Program Files\Notepad++\notepad++.exe"));
                    //執行
                    TaskService.Instance.RootFolder.RegisterTaskDefinition(jobName, td).Run();
                    
                    break;
               
                //修改工作排程
                case "UpdateTask":
                    Task updatetask = ts.GetTask(jobName);
                    if (updatetask == null) return;
                    updatetask.Definition.Triggers[0].StartBoundary = DateTime.Today + TimeSpan.FromDays(7);
                    updatetask.RegisterChanges();
                    break;

                //讀取工作排程
                case "ReadTask":
                    using (TaskService taskService = new TaskService())
                    {
                        foreach (Task task in taskService.AllTasks)
                        {
                            Console.WriteLine($"Task Name: {task.Name}");
                            // You can access other properties of the task as needed
                        }
                    }
                    //List<Microsoft.Win32.TaskScheduler.Task> getalltasks = ts.AllTasks;
                    //Task readtask = ts.GetTask(jobName);
                    
                    break;


                //刪除工作排程
                case "DeleteTask":

                    Task deletetask = ts.GetTask(jobName);
                    if (deletetask == null) return;
                    ts.RootFolder.DeleteTask(jobName);

                    break;

            }

        }

    }

}
