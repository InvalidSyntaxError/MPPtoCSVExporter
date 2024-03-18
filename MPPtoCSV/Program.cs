using System;
using System.IO;
using System.Text;
using net.sf.mpxj;
using net.sf.mpxj.MpxjUtilities;
using net.sf.mpxj.reader;

namespace MpxjQuery {
    class MpxjQuery {

        /// <summary>
        /// Main entry point.
        /// </summary>
        /// <param name="args">command line arguments</param>
        static void Main(string[] args) {
            bool bShowHelp=false;
            string strError="";
            string fileIn="", fileOut="";

            try {
                if (args.Length == 0) {
                    bShowHelp = true;
                    //System.Console.WriteLine("Usage: MpxQuery <input file name>");
                    //query("D:\\Projects\\C#\\MPP Viewer Source\\Microsoft-Project-Example.mpp");
                } else {
                    foreach (string arg in args) {
                        if (arg.ToLower() == "/help" ||
                            arg.ToLower() == "-?" ||
                            arg.ToLower() == "/?") {
                            bShowHelp = true;
                        } else if (arg.ToLower().Substring(0, 1) == "/") {
                            if (arg.ToLower().Substring(1, 1) == "o") {
                            } else {
                                strError += "unknown argument: " + arg + "\r\n";
                            }

                        } else {
                            if (fileIn == "") {
                                fileIn = arg;
                            } else {
                                fileOut = arg;
                            }
                        }
                    }
                }
                if (bShowHelp) {
                    PrintHelp();
                } else {
                    if (strError != "") {
                        strError += "use /? for help \r\n";
                        throw new ArgumentException(strError);
                    }
                }
                exportToCsv(fileIn, fileOut);
            } catch (Exception ex) {
                System.Console.WriteLine(ex.StackTrace);
            }

        }
        static void PrintHelp() {
            Console.WriteLine("Extract data from MS-Project files and exports them as CSV.");
            Console.WriteLine("       JKubik 2024");
            Console.WriteLine();
            Console.WriteLine("mpptocsv [/?] [/Od] source destination");
            Console.WriteLine();
            Console.WriteLine("source         source filepath to load. See mpxj what fileformats are supported.");
            Console.WriteLine("destination    output filepath");
            Console.WriteLine("/Od            tbd");
            Console.WriteLine();
            Console.WriteLine("Overwrites existing files.");
            Console.WriteLine("press any key...");
            Console.ReadKey();
        }

        static ProjectFile mpx;
        static StreamWriter m_LogFile;
        static string DELIM = "\t";
        private static void exportToCsv(string fileIn, string fileOut) {
            ProjectReader reader = ProjectReaderUtility.getProjectReader(fileIn);
            mpx = reader.read(fileIn); //load file, this will throw error if file doesnt exist

            //check for out-directory
            // if (!Directory.Exists(fileout)) {
            //    throw (new DirectoryNotFoundException(fileout));
            //} else {
            File.Delete(fileOut);
            m_LogFile = new StreamWriter(File.Create(fileOut),Encoding.ASCII,1000000);
            //}

            //export resource
            m_LogFile.WriteLine("Resources");
            m_LogFile.WriteLine("UniqueID" + DELIM + "ID" + DELIM + "Name");
            foreach (Resource resource in mpx.Resources.ToIEnumerable()) {
                m_LogFile.WriteLine(resource.UniqueID+DELIM+resource.ID+DELIM+resource.Name);
            }

            //export Tasks
            m_LogFile.WriteLine();
            m_LogFile.WriteLine("Tasks");
            m_LogFile.WriteLine("UniqueID" + DELIM + "ID" + DELIM + "Name"+DELIM+"Start"+DELIM+"End"+DELIM+"Duration");
            foreach (Task task in mpx.Tasks.ToIEnumerable()) {
                String startDate;
                String finishDate;
                String duration;
                Duration dur;

                var date = task.Start;
                if (date != null) {
                    startDate = date.ToDateTime().ToString();
                } else {
                    startDate = "";
                }

                date = task.Finish;
                if (date != null) {
                    finishDate = date.ToDateTime().ToString();
                } else {
                    finishDate = "";
                }

                dur = task.Duration;
                if (dur != null) {
                    duration = dur.toString();
                } else {
                    duration = "";
                }

                String baselineDuration = task.BaselineDurationText;
                if (baselineDuration == null) {
                    dur = task.BaselineDuration;
                    if (dur != null) {
                        baselineDuration = dur.toString();
                    } else {
                        baselineDuration = "";
                    }
                }
                m_LogFile.WriteLine(task.UniqueID + DELIM + task.ID + DELIM + task.Name + DELIM + startDate + DELIM + finishDate + DELIM + duration);
            }

            //export Assignment
            String taskName;
            String resourceName;
            m_LogFile.WriteLine();
            m_LogFile.WriteLine("Assignments");
            m_LogFile.WriteLine("TaskName" + DELIM + "ResourceName");
            foreach (ResourceAssignment assignment in mpx.ResourceAssignments.ToIEnumerable()) {
                Task task;
                Resource resource;
                task = assignment.Task;
                if (task == null) {
                    taskName = "(null task)";
                } else {
                    taskName = task.Name;
                }

                resource = assignment.Resource;
                if (resource == null) {
                    resourceName = "(null resource)";
                } else {
                    resourceName = resource.Name;
                }
                m_LogFile.WriteLine(taskName + DELIM + resourceName);
                m_LogFile.Flush();
            }

            //export Relation
            //this creates a table with Task-ID in first Column, N Column Predecessors-Task-IDs, N Column Successors-Task-IDs, 
            int ListSize = 10;
            string Line= "TaskID" + DELIM;
            m_LogFile.WriteLine();
            m_LogFile.WriteLine(String.Format("Relations (max.{0})",ListSize));
            for(var i = 0; i < ListSize; i++) {
                Line += String.Format("Pre{0}", i + 1)+DELIM;
            }
            for (var i = 0; i < ListSize; i++) {
                Line += String.Format("Suc{0}", i + 1) + DELIM;
            }
            m_LogFile.WriteLine(Line);
            
            foreach (Task task in mpx.Tasks.ToIEnumerable()) {
                java.util.List ListA = new java.util.ArrayList(ListSize);
                for (var i=0;i< ListSize; i++) {
                    if (task.Predecessors.size() > i) {
                        ListA.add(task.Predecessors.get(i));
                    } else {
                        ListA.add(null);
                    }
                }
                for (var i = 0; i < ListSize; i++) {
                    if (task.Successors.size() > i) {
                        ListA.add(task.Successors.get(i));
                    } else {
                        ListA.add(null);
                    }
                }
                Line = task.UniqueID + DELIM;
                for(var i=0; i<2*ListSize;i++) {
                    string subTaskID = "";
                    if(ListA.get(i)!=null) {
                        Relation Rel = ((Relation)ListA.get(i));
                        //if (i < ListSize) {
                            subTaskID= Rel.TargetTask.UniqueID.ToString();
                        //} else {
                        //    subTaskID = Rel.SourceTask.UniqueID.ToString();
                        //}
                    }
                    Line += subTaskID + DELIM;
                }
                m_LogFile.WriteLine(Line);
                m_LogFile.Flush();

            }
            m_LogFile.Flush();
            m_LogFile.Close();
        }

        /// <summary>
        /// This method performs a set of queries to retrieve information
        /// from the an MPP or an MPX file.
        /// </summary>
        /// <param name="filename">name of the project file</param>
        private static void query(String filename) {
            ProjectReader reader = ProjectReaderUtility.getProjectReader(filename);
            ProjectFile mpx = reader.read(filename);

            System.Console.WriteLine("MPP file type: " + mpx.ProjectProperties.MppFileType);

            listProjectHeader(mpx);

            listResources(mpx);

            listTasks(mpx);

            listAssignments(mpx);

            listAssignmentsByTask(mpx);

            listAssignmentsByResource(mpx);

            listHierarchy(mpx);

            listTaskNotes(mpx);

            listResourceNotes(mpx);

            listRelationships(mpx);

            listSlack(mpx);

            listCalendars(mpx);

            listCustomFields(mpx);
        }

        /// <summary>
        /// Reads basic summary details from the project header.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listProjectHeader(ProjectFile file) {
            ProjectProperties header = file.ProjectProperties;
            String formattedStartDate = header.StartDate == null ? "(none)" : header.StartDate.ToDateTime().ToString();
            String formattedFinishDate = header.FinishDate == null ? "(none)" : header.FinishDate.ToDateTime().ToString();

            System.Console.WriteLine("Project Header: StartDate=" + formattedStartDate + " FinishDate=" + formattedFinishDate);
            System.Console.WriteLine();
        }

        /// <summary>
        /// This method lists all resources defined in the file.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listResources(ProjectFile file) {
            foreach (Resource resource in file.Resources.ToIEnumerable()) {
                System.Console.WriteLine("Resource: " + resource.Name + " (Unique ID=" + resource.UniqueID + ") Start=" + resource.Start + " Finish=" + resource.Finish);
            }
            System.Console.WriteLine();
        }

        /// <summary>
        /// This method lists all tasks defined in the file.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listTasks(ProjectFile file) {
            foreach (Task task in file.Tasks.ToIEnumerable()) {
                String startDate;
                String finishDate;
                String duration;
                Duration dur;

                var date = task.Start;
                if (date != null) {
                    startDate = date.ToDateTime().ToString();
                } else {
                    startDate = "(no date supplied)";
                }

                date = task.Finish;
                if (date != null) {
                    finishDate = date.ToDateTime().ToString();
                } else {
                    finishDate = "(no date supplied)";
                }

                dur = task.Duration;
                if (dur != null) {
                    duration = dur.toString();
                } else {
                    duration = "(no duration supplied)";
                }

                String baselineDuration = task.BaselineDurationText;
                if (baselineDuration == null) {
                    dur = task.BaselineDuration;
                    if (dur != null) {
                        baselineDuration = dur.toString();
                    } else {
                        baselineDuration = "(no duration supplied)";
                    }
                }

                System.Console.WriteLine("Task: " + task.Name + " ID=" + task.ID + " Unique ID=" + task.UniqueID + " (Start Date=" + startDate + " Finish Date=" + finishDate + " Duration=" + duration + " Baseline Duration=" + baselineDuration + " Outline Level=" + task.OutlineLevel + " Outline Number=" + task.OutlineNumber + " Recurring=" + task.Recurring + ")");
            }
            System.Console.WriteLine();
        }

        /// <summary>
        /// This method lists all tasks defined in the file in a hierarchical format, 
        /// reflecting the parent-child relationships between them.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listHierarchy(ProjectFile file) {
            foreach (Task task in file.ChildTasks.ToIEnumerable()) {
                System.Console.WriteLine("Task: " + task.Name);
                listHierarchy(task, " ");
            }

            System.Console.WriteLine();
        }

        /// <summary>
        /// Helper method called recursively to list child tasks.
        /// </summary>
        /// <param name="task">Task instance</param>
        /// <param name="indent">print indent</param>
        private static void listHierarchy(Task task, String indent) {
            foreach (Task child in task.ChildTasks.ToIEnumerable()) {
                System.Console.WriteLine(indent + "Task: " + child.Name);
                listHierarchy(child, indent + " ");
            }
        }

        /// <summary>
        /// This method lists all resource assignments defined in the file.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listAssignments(ProjectFile file) {
            Task task;
            Resource resource;
            String taskName;
            String resourceName;

            foreach (ResourceAssignment assignment in file.ResourceAssignments.ToIEnumerable()) {
                task = assignment.Task;
                if (task == null) {
                    taskName = "(null task)";
                } else {
                    taskName = task.Name;
                }

                resource = assignment.Resource;
                if (resource == null) {
                    resourceName = "(null resource)";
                } else {
                    resourceName = resource.Name;
                }

                System.Console.WriteLine("Assignment: Task=" + taskName + " Resource=" + resourceName);
            }

            System.Console.WriteLine();
        }

        /// <summary>
        /// This method displays the resource assignments for each task. 
        /// This time rather than just iterating through the list of all 
        /// assignments in the file, we extract the assignments on a task-by-task basis.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listAssignmentsByTask(ProjectFile file) {
            foreach (Task task in file.Tasks.ToIEnumerable()) {
                System.Console.WriteLine("Assignments for task " + task.Name + ":");

                foreach (ResourceAssignment assignment in task.ResourceAssignments.ToIEnumerable()) {
                    Resource resource = assignment.Resource;
                    String resourceName;

                    if (resource == null) {
                        resourceName = "(null resource)";
                    } else {
                        resourceName = resource.Name;
                    }

                    System.Console.WriteLine("   " + resourceName);
                }
            }

            System.Console.WriteLine();
        }

        /// <summary>
        /// This method displays the resource assignments for each resource. 
        /// This time rather than just iterating through the list of all 
        /// assignments in the file, we extract the assignments on a resource-by-resource basis.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listAssignmentsByResource(ProjectFile file) {
            foreach (Resource resource in file.Resources.ToIEnumerable()) {
                System.Console.WriteLine("Assignments for resource " + resource.Name + ":");

                foreach (ResourceAssignment assignment in resource.TaskAssignments.ToIEnumerable()) {
                    Task task = assignment.Task;
                    System.Console.WriteLine("   " + task.Name);
                }
            }

            System.Console.WriteLine();
        }

        /// <summary>
        /// This method lists any notes attached to tasks..
        /// </summary>
        /// <param name="file">project file</param>
        private static void listTaskNotes(ProjectFile file) {
            foreach (Task task in file.Tasks.ToIEnumerable()) {
                String notes = task.Notes;

                if (notes != null && notes.Length != 0) {
                    System.Console.WriteLine("Notes for " + task.Name + ": " + notes);
                }
            }

            System.Console.WriteLine();
        }

        /// <summary>
        /// This method lists any notes attached to resources.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listResourceNotes(ProjectFile file) {
            foreach (Resource resource in file.Resources.ToIEnumerable()) {
                String notes = resource.Notes;

                if (notes != null && notes.Length != 0) {
                    System.Console.WriteLine("Notes for " + resource.Name + ": " + notes);
                }
            }

            System.Console.WriteLine();
        }

        /// <summary>
        /// This method lists task predecessor and successor relationships.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listRelationships(ProjectFile file) {
            foreach (Task task in file.Tasks.ToIEnumerable()) {
                System.Console.Write(task.ID);
                System.Console.Write('\t');
                System.Console.Write(task.Name);
                System.Console.Write('\t');

                dumpRelationList(task.Predecessors);
                System.Console.Write('\t');
                dumpRelationList(task.Successors);
                System.Console.WriteLine();
            }
        }

        /// <summary>
        /// Internal utility to dump relationship lists in a structured format that can 
        /// easily be compared with the tabular data in MS Project.
        /// </summary>
        /// <param name="relations">project file</param>
        private static void dumpRelationList(java.util.List relations) {
            if (relations != null && relations.isEmpty() == false) {
                if (relations.size() > 1) {
                    System.Console.Write('"');
                }
                bool first = true;
                foreach (Relation relation in relations.ToIEnumerable()) {
                    if (!first) {
                        System.Console.Write(',');
                    }
                    first = false;
                    System.Console.Write(relation.TargetTask.ID);
                    Duration lag = relation.Lag;
                    if (relation.Type != RelationType.FINISH_START || lag.Duration != 0) {
                        System.Console.Write(relation.Type);
                    }

                    if (lag.Duration != 0) {
                        if (lag.Duration > 0) {
                            System.Console.Write("+");
                        }
                        System.Console.Write(lag);
                    }
                }
                if (relations.size() > 1) {
                    System.Console.Write('"');
                }
            }
        }

        /// <summary>
        /// List the slack values for each task.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listSlack(ProjectFile file) {
            foreach (Task task in file.Tasks.ToIEnumerable()) {
                System.Console.WriteLine(task.Name + " Total Slack=" + task.TotalSlack + " Start Slack=" + task.StartSlack + " Finish Slack=" + task.FinishSlack);
            }
        }

        /// <summary>
        /// List details of all calendars in the file.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listCalendars(ProjectFile file) {
            foreach (ProjectCalendar cal in file.Calendars.ToIEnumerable()) {
                System.Console.WriteLine(cal.toString());
            }
        }

        /// <summary>
        /// List details of custom fields in the file.
        /// </summary>
        /// <param name="file">project file</param>
        private static void listCustomFields(ProjectFile file) {
            foreach (CustomField field in file.CustomFields.ToIEnumerable()) {
                System.Console.WriteLine(field.toString());
            }
        }
    }
}