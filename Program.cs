using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

namespace AposeWordThreadingIssue
{
    class Program
    {
        private static SemaphoreSlim semaphore;
        
        public static void Main()
        {
            // Create the semaphore.
            var threadsToRunAtOnce = 10;
            semaphore = new SemaphoreSlim(0, threadsToRunAtOnce);
            Console.WriteLine("{0} tasks can enter the semaphore.",
                semaphore.CurrentCount);
            var tasks = new Task[threadsToRunAtOnce];

            // Create and start five numbered tasks.
            for (int i = 0; i < threadsToRunAtOnce; i++)
            {
                tasks[i] = Task.Run(() =>
                {
                    var doc = new Document("Doc1.docx");

                    var names = new [] {"Title", "First_Name"};
                    var data = new object[] {"Title", "First_Name"};
                    
                    semaphore.Wait();
                    doc.MailMerge.Execute(names, data);
                });
            }

            // Wait to allow all the tasks to start and block.
            Thread.Sleep(2000);
            

            // Restore the semaphore count to its maximum value.
            Console.Write("Main thread calls Release() to start the work");
            semaphore.Release(threadsToRunAtOnce);
            Console.WriteLine("{0} tasks can enter the semaphore.",
                semaphore.CurrentCount);
            // Main thread waits for the tasks to complete.
            Task.WaitAll(tasks);

            Console.WriteLine("Main thread exits.");
        }
    }
}