using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.IO;

namespace Office365ExchangeTools
{
    class ExportPSTTool
    {
        private static OutlookApplication oa;

        static void Main(string[] args)
        {
            //load Args into arghandler
            Args argHandler = new Args(args);
            
            //opening Text, if this is stuck it failed to find the outlook application on the following line     
            Console.Write("Connecting to Outlook...");
            oa = new OutlookApplication();
            
            //Outlook found, printing version
            Console.Write("Found Version " + oa.Version + "\r\n");

            //List threading info
            GetTheadInfo();

            //List Outlook account status
            oa.ListAccounts();

            //Start console
            Console.WriteLine("type help for available commands.\r\n");

            //if their are arguements start the export process automatically
            if (args.Length > 0)
            {
                oa.ExportMailbox();
            }
            else
            {
                string command = "";
                while (command != "exit")
                {
                    Console.Write("ExportPSTTool: ");
                    command = Console.ReadLine();

                    switch (command.ToLower())
                    {
                        case "export":
                            oa.ExportMailbox();
                            break;
                        case "list":
                            oa.ListAccounts();
                            break;
                        case "help":
                            PrintHelp("Help");
                            break;
                        case "remove":
                            PrintHelp("remove");
                            break;
                        case "remove store":
                            oa.RemoveStore();
                            break;
                        case "threadinfo":
                            GetTheadInfo();
                            break;
                        case "setthreads":
                            SetTheadInfo();
                            break;
                        default:
                            break;
                    }
                }
            }
            Environment.Exit(0);
        }
        static void PrintHelp(string helpfile)
        {
            Console.WriteLine(Properties.Resources.ResourceManager.GetObject(helpfile));           
        }
        static void GetTheadInfo()
        {
            //get threading settings
            int minWorker;
            int minIOC;
            ThreadPool.GetMinThreads(out minWorker, out minIOC);

            Console.WriteLine("Current Minimum Worker Threads are : " + minWorker);
            Console.WriteLine("Current Minimum IO Completion: " + minIOC);
        }
        static void SetTheadInfo()
        {
            //get threading settings
            int minWorker;
            int minIOC;
            int newMinWorker = -1;
            int newMinIOC = -1;
            ThreadPool.GetMinThreads(out minWorker, out minIOC);

            while (newMinWorker == -1 || newMinIOC == -1)
            {
                Console.Write("\r\nSet New Minimum Worker Threads: ");
                string strNewMinWorker = Console.ReadLine();

                if (strNewMinWorker == "")
                {
                    newMinWorker = minWorker;
                }
                else
                {
                    int.TryParse(strNewMinWorker, out newMinWorker);
                }

                Console.Write("\r\nCurrent Minimum IO Completion: ");
                string strNewMinIOC = Console.ReadLine();

                if (strNewMinIOC == "")
                {
                    newMinIOC = minIOC;
                }
                else
                {
                    int.TryParse(strNewMinIOC, out newMinIOC);
                }
            }

            try
            {
                ThreadPool.SetMinThreads(newMinWorker, newMinIOC);
            }
            catch(System.Exception e)
            {
                Console.WriteLine("The minimum number of threads was not changed.");
                Console.WriteLine(e);
            }
            finally
            {
                Console.WriteLine("Current Minimum Worker Threads are : " + minWorker);
                Console.WriteLine("Current Minimum IO Completion: " + minIOC);
                
            }
        }
    }

    class OutlookApplication
    {
        private Application _outlookApp;
        private NameSpace _nameSpace;
        private List<Task> _currentTasks;

        public OutlookApplication()
        {
            _outlookApp = new Application();
            _nameSpace = _outlookApp.GetNamespace("MAPI");
        }

        public void ExportMailbox() {

            ExportOptions exportOptions = new ExportOptions();
            _currentTasks = new List<Task>();

            //Mailbox Name and attach
            if (Args.ReturnArg("-mailBoxName") == null)
            {
                Console.Write("Please enter the mailbox you'd like to attach (UPN): ");
                exportOptions.mailBoxName = Console.ReadLine();
            }
            else
            {
                exportOptions.mailBoxName = Args.ReturnArg("-mailBoxName");
            }

            //Find users mailbox on exchange
            Recipient recipient;
            try
            {
                recipient = _nameSpace.CreateRecipient(exportOptions.mailBoxName);
                //Attempt to Attach
                recipient.Resolve();
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e);
                return;
            }

            //Defining Export Data File
            if (Args.ReturnArg("-exportPSTPath") == null)
            {
                bool writable = false;

                while (!writable)
                {
                    Console.Write("Please enter the PST Export path and filename: ");
                    exportOptions.exportPSTPath = Console.ReadLine();

                    try
                    {
                        writable = true;

                        //test write to file
                        if (File.Exists(exportOptions.exportPSTPath))
                        {
                            File.SetLastWriteTimeUtc(exportOptions.exportPSTPath, DateTime.UtcNow);
                        }
                    }
                    catch (System.Exception e)
                    {
                        Console.WriteLine(e);
                        writable = false;
                    }

                }

            }
            else
            {
                exportOptions.exportPSTPath = Args.ReturnArg("-exportPSTPath");

                if(exportOptions.exportPSTPath == "" || exportOptions.exportPSTPath == null)
                {
                    exportOptions.exportPSTPath = exportOptions.mailBoxName + ".pst";
                }
            }
            
            //Define Export Start Date
            if (Args.ReturnArg("-exportStart") == null)
            {
                
                while (exportOptions.exportStart == null)
                {
                    Console.Write("Filter Start Date (Blank for January 1, 0001): ");

                    string inputStart = Console.ReadLine();

                    if (inputStart == "")
                    {
                        exportOptions.exportStart = DateTime.MinValue;
                    }
                    else
                    {
                        try
                        {
                            exportOptions.exportStart = DateTime.Parse(inputStart);
                        }
                        catch (System.Exception e)
                        {
                            Console.WriteLine(e);
                            Console.WriteLine("Something bad happened. Try again.");
                        }
                    }
                }
            }
            else
            {
                exportOptions.exportStart = DateTime.Parse(Args.ReturnArg("-exportStart"));
            }

            //Define export End
            if (Args.ReturnArg("-exportEnd") == null)
            {
                while (exportOptions.exportEnd == null)
                {
                    Console.Write("Filter End Date (Blank for January 1, 10000): ");

                    string inputEnd = Console.ReadLine();

                    if (inputEnd == "")
                    {
                        exportOptions.exportEnd = DateTime.MaxValue;
                    }
                    else
                    {
                        try
                        {
                            exportOptions.exportEnd = DateTime.Parse(inputEnd);
                        }
                        catch (System.Exception e)
                        {
                            Console.WriteLine(e);
                            Console.WriteLine("Something bad happened. Try again.");
                        }
                    }
                }
            }
            else
            {
                exportOptions.exportEnd = DateTime.Parse(Args.ReturnArg("-exportEnd"));
            }


            //Define the mail item flags
            if(Args.ReturnArg("-mailItems") == null)
            {
                while (exportOptions.mailItemFlag == ExportFlag.NotSet)
                {
                    Console.Write("Mail Items - (F)ilter By Item Date Range. (E)xclude. (A)ll is default: ");
                    exportOptions.mailItemFlag = ExportOptions.ReadExportFlag(Console.ReadLine());
                }
            }
            else
            {
                exportOptions.mailItemFlag = ExportOptions.ReadExportFlag(Args.ReturnArg("-mailItems"));
            }

            //Define the appointment item flags
            if (Args.ReturnArg("-appointmentItems") == null)
            {

                while (exportOptions.appointmentItemFlag == ExportFlag.NotSet)
                {
                    Console.Write("Appointment Items - (F)ilter By Item Date Range. (E)xclude. (A)ll is default: ");                    
                    exportOptions.appointmentItemFlag = ExportOptions.ReadExportFlag(Console.ReadLine());
                }
            }
            else
            {
                exportOptions.appointmentItemFlag = ExportOptions.ReadExportFlag(Args.ReturnArg("-appointmentItems"));
            }

            //Define the meeting item flags
            if (Args.ReturnArg("-meetingItems") == null)
            {
                while (exportOptions.meetingItemFlag == ExportFlag.NotSet)
                {
                    Console.Write("Meeting Items - (F)ilter By Item Date Range. (E)xclude. (A)ll is default: ");
                    exportOptions.meetingItemFlag = ExportOptions.ReadExportFlag(Console.ReadLine());
                }
            }
            else
            {
                exportOptions.meetingItemFlag = ExportOptions.ReadExportFlag(Args.ReturnArg("-meetingItems"));
            }

            //Define the contact item flags
            if (Args.ReturnArg("-contactItems") == null)
            {
                while (exportOptions.contactItemFlag == ExportFlag.NotSet)
                {
                    Console.Write("Contact Items - (F)ilter By Item Date Range. (E)xclude. (A)ll is default: ");
                    exportOptions.contactItemFlag = ExportOptions.ReadExportFlag(Console.ReadLine());
                }
            }
            else
            {
                exportOptions.contactItemFlag = ExportOptions.ReadExportFlag(Args.ReturnArg("-contactItems"));
            }

            //Define the contact item flags
            if (Args.ReturnArg("-taskItems") == null)
            {
                while (exportOptions.taskItemFlag == ExportFlag.NotSet)
                {
                    Console.Write("Task Items - (F)ilter By Item Date Range. (E)xclude. (A)ll is default: ");
                    exportOptions.taskItemFlag = ExportOptions.ReadExportFlag(Console.ReadLine());
                }
            }
            else
            {
                exportOptions.taskItemFlag = ExportOptions.ReadExportFlag(Args.ReturnArg("-taskItems"));
            }

            //Define the contact item flags
            if (Args.ReturnArg("-journalItems") == null)
            {
                while (exportOptions.journalItemFlag == ExportFlag.NotSet)
                {
                    Console.Write("Journal Items - (F)ilter By Item Date Range. (E)xclude. (A)ll is default: ");
                    exportOptions.journalItemFlag = ExportOptions.ReadExportFlag(Console.ReadLine());
                }
            }
            else
            {
                exportOptions.journalItemFlag = ExportOptions.ReadExportFlag(Args.ReturnArg("-taskItems"));
            }

            //Define the contact item flags
            if (Args.ReturnArg("-otherItems") == null)
            {
                while (exportOptions.otherItemFlag == ExportFlag.NotSet)
                {
                    Console.Write("All other Items - (F)ilter By Item Date Range. (E)xclude. (A)ll is default: ");
                    exportOptions.otherItemFlag = ExportOptions.ReadExportFlag(Console.ReadLine());
                }
            }
            else
            {
                exportOptions.otherItemFlag = ExportOptions.ReadExportFlag(Args.ReturnArg("-otherItems"));
            }


            Folder backupPSTRootFolder;
            //Add a new store in the pst path.

            try
            {
                _nameSpace.Session.AddStoreEx(exportOptions.exportPSTPath.ToString(), OlStoreType.olStoreUnicode);
            }
            catch(System.Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("Unable to create backup pst. Aborting...");
            }
            finally
            {
                //find the folder created with the new store
                foreach (Folder folder in _nameSpace.Session.Folders)
                {
                    if (folder.Store.FilePath != null && folder.Store.FilePath.ToLower() == exportOptions.exportPSTPath.ToLower())
                    {
                        backupPSTRootFolder = folder;

                        foreach (Folder accountFolder in _nameSpace.Folders)
                        {
                            //need to find the user account amongst my own email folders
                            if (accountFolder.Name.Contains(recipient.Name))
                            {
                                string file = Path.GetDirectoryName(exportOptions.exportPSTPath.ToString());
                                DirectoryInfo root = Directory.CreateDirectory(file + "\\" + recipient.Name);

                                //RecurseFolders(backupPSTRootFolder, accountFolder.Folders, exportOptions, root);
                                Console.WriteLine("Scheduling recursion of root folder " + accountFolder.FullFolderPath);
                                RecurseFolders(backupPSTRootFolder, accountFolder.Folders, exportOptions, root);
                            }
                        }
                    }
                }
            }

            Task.WaitAll(_currentTasks.ToArray<Task>());

            Console.WriteLine("Export Ended.");
        }
        private async void RecurseFolders(Folder backupFolder,Folders folders, ExportOptions exportOptions, DirectoryInfo root)
        {
            foreach(Folder folder in backupFolder.Folders)
            {
                if(folder.Name != "Deleted Items" && 
                   folder.Name != "Search Folders" && 
                   folder.Name != "PersonMetadata" &&
                   folder.Name != "Recipient Cache" &&
                   folder.Name != "Sync Issues" &&
                   folder.Name != "Yammer Root") 
                {
                    folder.Delete();
                }
            }

            foreach (Folder folder in folders)
            {
                try
                {
                    if (
                        (!folder.IsSharePointFolder) &&
                        (folder.Name != "Deleted Items" &&
                         folder.Name != "Search Folders" &&
                         folder.Name != "PersonMetadata" &&
                         folder.Name != "Recipient Cache" &&
                         folder.Name != "Sync Issues" &&
                         folder.Name != "Yammer Root" &&
                         folder.Name != "Junk Email" &&
                         folder.Name != "Junk E-mail" &&
                         folder.Name != "Clutter") &&
                        (folder.Name == "Calendar" && exportOptions.meetingItemFlag != ExportFlag.Exclude) ||
                        (folder.Name == "Calendar" && exportOptions.appointmentItemFlag != ExportFlag.Exclude) ||
                        (folder.Name == "Journal" && exportOptions.journalItemFlag != ExportFlag.Exclude) ||
                        (folder.Name == "Inbox" && exportOptions.mailItemFlag != ExportFlag.Exclude) ||
                        (folder.Name == "Tasks" && exportOptions.taskItemFlag != ExportFlag.Exclude)
                       )
                    {
                        MAPIFolder newbackupfolder;
                        
                        switch (folder.DefaultMessageClass)
                        {
                            case "IPM.Appointment":
                                newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderCalendar);
                                break;
                            case "IPM.Task":
                                newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderTasks);
                                break;
                            case "IPM.Activity":
                                newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderJournal);
                                break;
                            case "IPM.Contact":
                                newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderContacts);
                                break;
                            case "IPM.Post":
                                newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderNotes);
                                break;
                            case "IPM.StickyNote":
                                newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderNotes);
                                break;
                            default:
                                switch (folder.Name)
                                {
                                    case "Inbox":
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderInbox);
                                        break;
                                    case "Outbox":
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderOutbox);
                                        break;
                                    case "Sent Items":
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderSentMail);
                                        break;
                                    case "Junk Email":
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderJunk);
                                        break;
                                    case "Junk E-mail":
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderJunk);
                                        break;
                                    case "Drafts":
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderDrafts);
                                        break;
                                    case "RSS Feeds":
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name, OlDefaultFolders.olFolderRssFeeds);
                                        break;
                                    default:
                                        newbackupfolder = backupFolder.Folders.Add(folder.Name);
                                        break;
                                }                                
                                break;
                        
                        }

                        root = Directory.CreateDirectory(root.FullName + "\\" + folder.Name);

                        Console.WriteLine("Scheduling recursion of folder " + folder.FullFolderPath);
                        RecurseFolders((Folder)newbackupfolder, folder.Folders, exportOptions, root);

                        _currentTasks.Add(Task.Run(async () =>
                            {
                                Console.WriteLine("Scheduling backup of folder " + newbackupfolder.FullFolderPath);
                                await BackupItems((Folder)newbackupfolder, folder, exportOptions,root);
                            })
                        );

                        //BackupItems((Folder)newbackupfolder, folder, exportOptions, root);

                    }
                                       
                }
                catch(System.Exception e)
                {
                    Console.WriteLine(e);
                }
                
            }

            Task.WaitAll(_currentTasks.ToArray<Task>());


        }
        private async Task BackupItems(Folder backupFolder,Folder folder, ExportOptions exportOptions, DirectoryInfo root)
        {
            await Task.Yield();

            int i = 1;
            foreach (object item in folder.Items)
            {

                _currentTasks.Add(Task.Run(async () =>
                    {
                        Console.WriteLine("Scheduling CopyMove on " + backupFolder.FullFolderPath + " Item # " + i);
                        await CopyMove(item, backupFolder, exportOptions, root);
                        Marshal.ReleaseComObject(item);  //cannot set to null before this call
                        Marshal.FinalReleaseComObject(item);
                    })
                );
                
                i++;
            }

            

        }
        private async Task CopyMove(object item, Folder backupFolder,ExportOptions exportOptions, DirectoryInfo root)
        {
            await Task.Yield();
            if (item is MailItem)
            {
                try
                {
                    MailItem mailItem = (MailItem)item;
                    if (
                           (exportOptions.mailItemFlag == ExportFlag.All) ||
                           (exportOptions.mailItemFlag != ExportFlag.Exclude) ||
                           (exportOptions.mailItemFlag == ExportFlag.Filter && Between(mailItem.CreationTime, exportOptions.exportStart, exportOptions.exportEnd))
                       )
                    {
                        ////copy
                        //MailItem copiedMailItem = mailItem.Copy();
                        //copiedMailItem.Move(backupFolder);
                        //
                        ////audit
                        //Console.WriteLine(backupFolder.FolderPath + "\t MailItem \t" + mailItem.Subject + "\t" + mailItem.ReceivedTime);
                        //
                        //
                        ////close items
                        //mailItem.Close(OlInspectorClose.olDiscard);
                        //copiedMailItem.Close(OlInspectorClose.olDiscard);

                        mailItem.SaveAs(root.FullName + "\\" + mailItem.EntryID + ".msg", OlSaveAsType.olMSGUnicode);
                        mailItem.Close(OlInspectorClose.olDiscard);

                        Marshal.ReleaseComObject(mailItem);  //cannot set to null before this call
                        Marshal.FinalReleaseComObject(mailItem);

                    }
                   
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e);
                }
            }
            //else if (item is AppointmentItem)
            //{
            //    try
            //    {
            //        AppointmentItem appointmentItem = (AppointmentItem)item;
            //
            //        if ((exportOptions.appointmentItemFlag != ExportFlag.Exclude && Between(appointmentItem.CreationTime, exportOptions.exportStart, exportOptions.exportEnd)) || exportOptions.appointmentItemFlag == ExportFlag.All)
            //        {
            //            //copy
            //            appointmentItem.CopyTo((MAPIFolder)backupFolder, OlAppointmentCopyOptions.olCreateAppointment);
            //
            //            //audit
            //            Console.WriteLine(backupFolder.FolderPath + "\t AppointmentItem \t" + appointmentItem.Subject + "\t" + appointmentItem.StartUTC);
            //
            //            //closeitems
            //            appointmentItem.Close(OlInspectorClose.olDiscard);
            //        }
            //    }
            //    catch (System.Exception e)
            //    {
            //        Console.WriteLine(e);
            //    }
            //}
            //else if (item is ContactItem)
            //{
            //    try
            //    {
            //        ContactItem contactItem = (ContactItem)item;
            //
            //        if ((exportOptions.contactItemFlag != ExportFlag.Exclude && Between(contactItem.CreationTime, exportOptions.exportStart, exportOptions.exportEnd)) || exportOptions.contactItemFlag == ExportFlag.All)
            //        {
            //            //copy
            //            ContactItem copiedContactItem = contactItem.Copy();
            //            copiedContactItem.Move(backupFolder);
            //
            //            //audit
            //            Console.WriteLine("\t ContactItem \t" + copiedContactItem.LastName + ", " + copiedContactItem.FirstName);
            //            
            //            //closeitems
            //            contactItem.Close(OlInspectorClose.olDiscard);
            //            copiedContactItem.Close(OlInspectorClose.olDiscard);
            //        }
            //    }
            //    catch (System.Exception e)
            //    {
            //        Console.WriteLine(e);
            //    }
            //}
            //else if (item is MeetingItem)
            //{
            //    try
            //    {
            //        MeetingItem meetingItem = (MeetingItem)item;
            //
            //        if ((exportOptions.meetingItemFlag != ExportFlag.Exclude && Between(meetingItem.CreationTime, exportOptions.exportStart, exportOptions.exportEnd)) || exportOptions.meetingItemFlag == ExportFlag.All)
            //        {
            //            //copy
            //            MeetingItem copiedMeetingItem = meetingItem.Copy();
            //            copiedMeetingItem.Move(backupFolder);
            //
            //            //audit 
            //            Console.WriteLine(backupFolder.FolderPath + "\t MeetingItem \t" + copiedMeetingItem.Subject + "\t" + copiedMeetingItem.ReceivedTime);
            //            
            //            //closeitems
            //            meetingItem.Close(OlInspectorClose.olDiscard);
            //            copiedMeetingItem.Close(OlInspectorClose.olDiscard);
            //        }
            //    }
            //    catch (System.Exception e)
            //    {
            //        Console.WriteLine(e);
            //    }
            //}
            //else if (item is TaskItem)
            //{
            //    try
            //    {
            //        TaskItem taskItem = (TaskItem)item;
            //
            //        if ((exportOptions.taskItemFlag != ExportFlag.Exclude && Between(taskItem.CreationTime, exportOptions.exportStart, exportOptions.exportEnd)) || exportOptions.taskItemFlag == ExportFlag.All)
            //        {
            //
            //            //copy
            //            TaskItem copiedtaskItem = taskItem.Copy();
            //            copiedtaskItem.Move(backupFolder);
            //
            //            //audit
            //            Console.WriteLine(backupFolder.FolderPath + "\t TaskItem \t" + copiedtaskItem.Subject + "\t" + copiedtaskItem.CreationTime);
            //            
            //            //closeitems
            //            taskItem.Close(OlInspectorClose.olDiscard);
            //            copiedtaskItem.Close(OlInspectorClose.olDiscard);
            //        }
            //    }
            //    catch (System.Exception e)
            //    {
            //        Console.WriteLine(e);
            //    }
            //}
            //else if (item is JournalItem)
            //{
            //    try
            //    {
            //        JournalItem journalItem = (JournalItem)item;
            //
            //        if ((exportOptions.journalItemFlag != ExportFlag.Exclude && Between(journalItem.CreationTime, exportOptions.exportStart, exportOptions.exportEnd)) || exportOptions.journalItemFlag == ExportFlag.All)
            //        {
            //            //copy
            //            JournalItem copiedjournalItem = journalItem.Copy();
            //            copiedjournalItem.Move(backupFolder);
            //
            //            //audit
            //            Console.WriteLine(backupFolder.FolderPath + "\t JournalItem \t" + copiedjournalItem.Subject + "\t" + copiedjournalItem.CreationTime);
            //            
            //            //closeitems
            //            journalItem.Close(OlInspectorClose.olDiscard);
            //            copiedjournalItem.Close(OlInspectorClose.olDiscard);
            //        }
            //    }
            //    catch (System.Exception e)
            //    {
            //        Console.WriteLine(e);
            //    }
            //}
            //else
            //{
            //    try
            //    {
            //        if (exportOptions.otherItemFlag == ExportFlag.All)
            //        {
            //            Console.WriteLine("Couldn't Identify mail item at " + backupFolder.FolderPath);                  
            //        }
            //    }
            //    catch (System.Exception e)
            //    {
            //        Console.WriteLine(e);
            //    }
            //}
            
        }
        public void RemoveStore()
        {
            List<Store> stores = new List<Store>();

            foreach (Store store in _nameSpace.Stores)
            {
                stores.Add(store);

                Console.WriteLine("[" + stores.IndexOf(store) + "] " + store.DisplayName);
            }

            Console.Write("Type the number of the store you'd like to delete :");

            Store selectedStore = null;

            while(selectedStore == null)
            {
                string selectedValue = Console.ReadLine();
                int selectedKey;

                if(int.TryParse(selectedValue,out selectedKey))
                {
                    if (stores.ElementAt(selectedKey) !=null)
                    {
                        _nameSpace.RemoveStore(stores.ElementAt(selectedKey).GetRootFolder());
                    }
                }
            }

        }
        private static bool Between(DateTime input, DateTime start, DateTime end)
        {
            return (input > start && input < end);
        }
        private static bool Between(DateTime input, DateTime? start, DateTime? end)
        {
            return (input > start && input < end);
        }
        public string Version
        {
            get
            {
                return _outlookApp.Version;
            }
        }
        public void ListAccounts()
        {
            Console.WriteLine("\r\nThe following Accounts are found on this instance of Outlook:");

            int i = 1;

            foreach (Account account in _outlookApp.Session.Accounts)
            {
                Console.WriteLine("[" + i + "] " + account.DisplayName);
                i++;
            }

            Console.WriteLine("\r\nThe following Stores are attached to this instance of Outlook:");

            int j = 1;

            foreach (Store store in _nameSpace.Stores)
            {
                Console.WriteLine("[" + j + "] " + store.DisplayName);
                j++;
            }

            Console.WriteLine(" ");
        }

    }
    class Args
    {
        public string[] _args;
        public static Args _instance;

        public Args(string[] args)
        {
            _args = args;
            _instance = this;
        }

        public static string ReturnArg(string key)
        {
            if (Array.IndexOf(Args._instance._args, key) != -1 && Array.IndexOf(Args._instance._args, key) <= Args._instance._args.Length)
            {
                return Args._instance._args[Array.IndexOf(Args._instance._args, key) + 1];
            }
            else
            {
                return null;
            }
        }
    }
    class ExportOptions
    {
        public string mailBoxName;
        public string exportPSTPath;
        public DateTime? exportStart;
        public DateTime? exportEnd;
        public ExportFlag mailItemFlag;
        public ExportFlag appointmentItemFlag;
        public ExportFlag meetingItemFlag;
        public ExportFlag contactItemFlag;
        public ExportFlag taskItemFlag;
        public ExportFlag journalItemFlag;
        public ExportFlag otherItemFlag;
        public int maxCopyMoveAttempts;

        public ExportOptions()
        {
            this.exportStart = null;
            this.exportEnd = null;
            this.mailItemFlag = ExportFlag.NotSet;
            this.appointmentItemFlag = ExportFlag.NotSet;
            this.meetingItemFlag = ExportFlag.NotSet;
            this.contactItemFlag = ExportFlag.NotSet;
            this.taskItemFlag = ExportFlag.NotSet;
            this.journalItemFlag = ExportFlag.NotSet;
            this.otherItemFlag = ExportFlag.NotSet;
            this.maxCopyMoveAttempts = 5;
        }

        public static ExportFlag ReadExportFlag(string flag)
        {
            ExportFlag returnFlag = ExportFlag.NotSet;

            switch (flag.ToLower())
            {
                case "e":
                    returnFlag = ExportFlag.Exclude;
                    break;
                case "f":
                    returnFlag = ExportFlag.Filter;
                    break;
                case "a":
                case "":
                    returnFlag = ExportFlag.All;
                    break;
            }

            return returnFlag;
        }
    }
    enum ExportFlag
    {
        Filter,
        Exclude,
        All,
        NotSet
    }
}
