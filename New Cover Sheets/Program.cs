using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using IWshRuntimeLibrary;
using Excel = Microsoft.Office.Interop.Excel;
using File = System.IO.File;
using System.Globalization;
//This program creates the pw structure for our initial vetting
//CA folders have descriptions
//Facility/Component folders have numbers
//Drawings have numbers, but those in the drawings folder have numbers and descriptions (they have a shorter path, so the fileName can be longer)
namespace Vetting_Folder_Structure
{
    public static class StringExt
    {
        public static string Truncate(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }
    class Facility
    {
        public List<string> Drawings;
        public string Name;
        public Facility(string name)
        {
            Name = name;
            Drawings = new List<string>();
        }
    }
    class Ca
    {
        public List<string> Drawings;
        public string Name;
        public Ca(string name)
        {
            Name = name;
            Drawings = new List<string>();
        }
    }
    class Program
    {
        public static string SpanInsert(string val)
        {
            if (val != null)
            {
                val = "<span>" + val.Replace("\n", "</span><br><span>") + "</span>";
                return val;
            }
            else
            {
                return null;
            }
        }
        public static string Capitalise(string input)
        {
            if (input == "")
            {
                return "";
            }
            input = input.ToLower().Trim();
            input =  input.First().ToString().ToUpper() +  input.Substring(1);
            return input;
        }

        public static string GetFacilityPdfString(SortedSet<string> pdfs)
        {
            string pdfString = "";
            string start = @"<div>
       <nav class=""panel"">
  <p class=""panel-heading"">
    Facility Files
  </p>
";
            string middle = "";
            if (pdfs.Count == 0)
            {
                middle += String.Format(@"
  <a class=""panel-block is-active"">
    <span class=""panel-icon"">
      <i class=""fa fa-book""></i>
    </span>
    No Facility Files
  </a>
");
            }
            foreach(var pdf in pdfs)
            {
                middle += String.Format(@"
  <a class=""panel-block is-active"">
    <span class=""panel-icon"">
      <i class=""fa fa-book""></i>
    </span>
    {0}
  </a>
",pdf);
            }
            string end = @"
  <div class=""panel-block"">
  </div>
</nav></div>
";
            pdfString = start + middle + end;
            return pdfString;
        }

        public static string GetCaPdfString(Dictionary<string,SortedSet<string>> caDict)
        {
            
            string start = @"<div>
       <nav class=""panel"">
  <p class=""panel-heading"">
    Construction Activity Files
  </p>
";
            string pdfString = start;
            foreach(KeyValuePair<string,SortedSet<string>> ca in caDict)
            {
                string middlePart = "";
                string newStart = String.Format(@"
       <nav class=""panel"">
  <p class=""panel-heading"">
    Construction Activity Number: {0}
  </p>
",ca.Key);

                string middle = "";
                if (ca.Value.Count == 0)
                {
                    middle += String.Format(@"
  <a class=""panel-block is-active"">
    <span class=""panel-icon"">
      <i class=""fa fa-book""></i>
    </span>
    No Construction Activity Files
  </a>
");
                }
                foreach (var pdf in ca.Value)
                {
                    middle += String.Format(@"
  <a class=""panel-block is-active"">
    <span class=""panel-icon"">
      <i class=""fa fa-book""></i>
    </span>
    {0}
  </a>
", pdf);
                }
                middlePart = newStart + middle;
                pdfString += middlePart;
            }
            string end = @"
  <div class=""panel-block"">
  </div>
</nav></div>
";
            pdfString += end;
            return pdfString;
        }
        static void Main(string[] args)
        {
            string scoeBaseFolder = "C:\\SCoE";
            if (!Directory.Exists(scoeBaseFolder)) //create folder if it does not exist
            {
                Directory.CreateDirectory(scoeBaseFolder);
            }
            if (!Directory.Exists(Path.Combine(scoeBaseFolder, "Drawings")))
            {
                Directory.CreateDirectory(Path.Combine(scoeBaseFolder, "Drawings"));
            }
            DirectoryInfo drawDir = new DirectoryInfo(Path.Combine(scoeBaseFolder, "Drawings"));

            string vFilePath = @"C:\Users\CCrowe\IdeaProjects\GetFacilityDetails\data.xlsx";
                //before we used C:\\Users\\ccrowe\\Desktop\\JCMS Image\\Copy of Appendix A SCoE Facility List.xlsx 
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            Excel.Workbook vWb = xl.Workbooks.Open(vFilePath);
            Excel.Worksheet vWs = vWb.Sheets["Datas"];

            //string checkFile = "C:\\Users\\ccrowe\\Documents\\Appendix A SCoE Facility List2.xlsx";
                //Added later, check this list to see if we want to create the cover sheet
            //Excel.Workbook cWb = xl.Workbooks.Open(checkFile);
            //Excel.Worksheet cs1 = cWb.Sheets[1];

            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;


            using (
                SqlConnection conn =
                    new SqlConnection("Server=OME-CND6435DR5;Database=JCMS_Local;Integrated Security = true"))
            {

                conn.Open();
                string sql;
                bool foundNum = true;
                int lastrow = vWs.UsedRange.Rows.Count;
                for (int i = 2; i <= vWs.UsedRange.Rows.Count; i++) //vWs.UsedRange.Rows.Count  !!!!!Make sure these rows are correct!
                {
                    string fname = vWs.Range["B" + i.ToString()].Value;
                    if (fname != null)
                    {
                        /*bool foundNum = false;
                    for (int j = 3; j <= cs1.UsedRange.Rows.Count; j++)
                    {
                        if (cs1.Range["A" + j.ToString()].Value == fname)
                        {
                            foundNum = true;
                            break;
                        }
                    }*/
                        if (foundNum == true)
                        {

                            string secondaryProponent = vWs.Range["P" + i.ToString()].Value;
                            if (secondaryProponent != null && secondaryProponent != "")
                            {
                                secondaryProponent = secondaryProponent.Trim();
                            }
                            else
                            {
                                secondaryProponent = "No Secondary Proponent";
                            }

                            if (!Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent)))
                            {
                                Directory.CreateDirectory(Path.Combine(scoeBaseFolder, secondaryProponent));
                            }
                            string masterPlanningCategory = vWs.Range["J" + i.ToString()].Value;
                            if (masterPlanningCategory != null && masterPlanningCategory != "")
                            {
                                masterPlanningCategory = masterPlanningCategory.Trim();
                            }
                            else
                            {
                                masterPlanningCategory = "No Master Planning Category";
                            }
                            if (
                                !Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                    masterPlanningCategory)))
                            {
                                Directory.CreateDirectory(Path.Combine(scoeBaseFolder, secondaryProponent,
                                    masterPlanningCategory));
                            }
                            string fnameLong = fname;
                            if (
                                !Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                    masterPlanningCategory,
                                    fnameLong)))
                            {
                                Directory.CreateDirectory(Path.Combine(scoeBaseFolder, secondaryProponent,
                                    masterPlanningCategory, fnameLong));
                            }
                            string fPath = Path.Combine(scoeBaseFolder, secondaryProponent, masterPlanningCategory,
                                fnameLong);
                            string facilityNumber = SpanInsert(vWs.Range["B" + i.ToString()].Value);
                            /*string newDescription = vWs.Range["D" + i.ToString()].Value;
                        if (newDescription != null)
                        {
                            newDescription = "<span>" + newDescription.Replace("\n", "</span><br><span>") + "</span>";
                        }*/
                            string origDescription = SpanInsert(vWs.Range["D" + i.ToString()].Value);
                            /*string designAgentComments = vWs.Range["N" + i.ToString()].Value;
                        if (designAgentComments != null)
                        {
                            designAgentComments = "<span>" + designAgentComments.Replace("\n", "</span><br><span>") + "</span>";
                        }*/
                            var r = new Regex(@"(^[a-z])|\.\s+(.)", RegexOptions.ExplicitCapture);

                            string newProponentComments = SpanInsert(Capitalise(vWs.Range["L" + i.ToString()].Value));


                            string designator = SpanInsert(Capitalise(vWs.Range["B" + i.ToString()].Value));//
                            string description = SpanInsert(Capitalise(vWs.Range["D" + i.ToString()].Value));//
                            string detailField = SpanInsert(Capitalise(vWs.Range["F" + i.ToString()].Value));//
                            string lookupToNoun = SpanInsert(Capitalise(vWs.Range["H" + i.ToString()].Value));//
                            string lookupToStandard = SpanInsert(Capitalise(vWs.Range["I" + i.ToString()].Value));//
                            string lookupToMasterPlanningCategory = SpanInsert(Capitalise(vWs.Range["J" + i.ToString()].Value));//
                            string primaryConstructionMaterial = SpanInsert(Capitalise(vWs.Range["E" + i.ToString()].Value));//
                            string primaryProponent = SpanInsert(Capitalise(vWs.Range["O" + i.ToString()].Value));//
                            string lookupToType = SpanInsert(Capitalise(vWs.Range["G" + i.ToString()].Value));//
                            string proponentRecommendation = SpanInsert(Capitalise(vWs.Range["K" + i.ToString()].Value));//
                            string vettingDate = SpanInsert(vWs.Range["Q" + i.ToString()].Value);

                            if (primaryProponent == "<span></span>")
                            {
                                primaryProponent = "No Primary Proponent";
                            }

                            if (vettingDate == "<span></span>")
                            {
                                vettingDate = "Has not been previously vetted";
                            }

                            Dictionary<string,SortedSet<string>> cas = new Dictionary<string,SortedSet<string>>();
                            SortedSet<string> facs = new SortedSet<string>();

                            string currentCA = null;
                            string currentCANumber = null;
                            //CONSTRUCTION ACTIVITY
                            if (lookupToType.Contains("Facility")) //Don't do this for components
                            {
                                sql = string.Format(
                                    @"DECLARE @ele_id nvarchar(50);
                        DECLARE @ele_name nvarchar(100);
                        DECLARE @ele_descr nvarchar(100);
                        DECLARE @ele_type nvarchar(100);
                        DECLARE @type nvarchar(50);
                        DECLARE @subtype nvarchar(50);
                        DECLARE @FetchStatus int
                        DECLARE CA_cursor CURSOR  
	                        FOR select Element_Id, Element_Nbr, Element_Descr, Element_Type FROM Element WHERE Element_Id in 
	                        (SELECT Element_Id FROM Element_Hierarchy WHERE Parent_Element_Id In 
	                        (SELECT Element_Id FROM Element WHERE Element_Nbr = '{0}')) ORDER BY Element_Nbr ASC;
	
                        OPEN CA_cursor  
                        FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name,@ele_descr, @ele_type
                        WHILE @@FETCH_STATUS = 0
	                        BEGIN
                                SELECT @ele_id, @ele_name, @ele_descr, @ele_type
                                SET @subtype = (SELECT Element_Type from Element where Element_Nbr = @ele_name);
		                        SELECT DISTINCT File_Link, File_Nbr,File_Title from JCMS_File WHERE File_Id in (SELECT File_Id FROM JCMS_File_Owner WHERE File_Owner_Id = @ele_id AND File_Owner_Obj_Type = @subtype) AND File_Class = 'DRAWING'  AND File_Type = 'PDF' ORDER BY File_Nbr ASC ;
		                        FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name, @ele_descr, @ele_type
	                        END
                        CLOSE CA_cursor;
                        DEALLOCATE CA_cursor;
                ", fname);
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    SqlDataReader reader = cmd.ExecuteReader();
                                    do
                                    {
                                        while (reader.Read())
                                        {
                                            string[] colNames = new string[4];
                                            reader.GetValues(colNames);
                                            var v = colNames[3];
                                            if (colNames[3] == null) //{string[0]} is null
                                            {
                                                string drawPath = colNames[0];
                                                string drawName = colNames[1] + " - " +
                                                                  colNames[2].Replace("/", "-")
                                                                      .Replace(":", "-")
                                                                      .Replace("\"", "-");
                                                drawName = drawName.Truncate(120);
                                                cas.Last().Value.Add(drawName);
                                                string shortDrawName = colNames[1];
                                                string finalDrawPath = drawDir.FullName + "\\" + drawName + ".pdf";
                                                try
                                                {
                                                    if (drawPath ==
                                                        "JCMS\\DATA_FILES\\DRAWINGS\\PDF\\70229-4108M-MP101X.PDF")
                                                    {
                                                        drawPath =
                                                            "JCMS\\DATA_FILES\\DRAWINGS\\PDF\\70229-4108_M-MP101X.PDF";
                                                        //fix, wrong in the database
                                                    }
                                                    File.Copy("C:" + "\\" + drawPath, finalDrawPath, true);
                                                }
                                                catch (Exception excpt)
                                                {
                                                    Console.WriteLine("Error with file..."); //70229-4108_M-MP101X
                                                    Console.WriteLine(drawName);
                                                    Console.WriteLine();
                                                }
                                                //Now add a shortcut in that file
                                                string finalLinkPath = null;
                                                if (
                                                    !Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "CA Drawings")))
                                                {
                                                    Directory.CreateDirectory(Path.Combine(scoeBaseFolder,
                                                        secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "CA Drawings"));
                                                }
                                                if (
                                                    !Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                                        masterPlanningCategory,
                                                        fnameLong, "CA Drawings",
                                                        currentCA.Replace("/", "-").Replace(":", "-").Replace("\"", "-"))))
                                                {
                                                    Directory.Move(
                                                        Path.Combine(scoeBaseFolder, secondaryProponent,
                                                            masterPlanningCategory,
                                                            fnameLong, "CA Drawings",
                                                            currentCA.Replace("/", "-")
                                                                .Replace(":", "-")
                                                                .Replace("\"", "-") +
                                                            "_Empty"),
                                                        Path.Combine(scoeBaseFolder, secondaryProponent,
                                                            masterPlanningCategory,
                                                            fnameLong, "CA Drawings",
                                                            currentCA.Replace("/", "-")
                                                                .Replace(":", "-")
                                                                .Replace("\"", "-")));
                                                }
                                                DirectoryInfo di =
                                                    new DirectoryInfo(Path.Combine(scoeBaseFolder, secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "CA Drawings"));
                                                string parent = "";
                                                foreach (
                                                    var file in
                                                        di.GetFiles("*" + currentCANumber + "*",
                                                            SearchOption.AllDirectories))
                                                {
                                                    parent = file.DirectoryName;
                                                }

                                                finalLinkPath = Path.Combine(parent,
                                                    shortDrawName + ".pdf.lnk");

                                                WshShell wsh = new WshShell();
                                                IWshRuntimeLibrary.IWshShortcut shortcut =
                                                    (IWshShortcut)
                                                        wsh.CreateShortcut(finalLinkPath);

                                                shortcut.Arguments = "";
                                                shortcut.TargetPath = finalDrawPath;
                                                bool isTargetPathValid = File.Exists(finalDrawPath);
                                                // not sure about what this is for
                                                shortcut.WindowStyle = 1;

                                                shortcut.Description = "Const Act:" +
                                                                       currentCA.Replace("/", "-")
                                                                           .Replace(":", "-")
                                                                           .Replace("\"", "-");
                                                shortcut.WorkingDirectory = drawDir.FullName;
                                                bool isWorkingValid = Directory.Exists(shortcut.WorkingDirectory);
                                                shortcut.IconLocation = "icon location";
                                                shortcut.Save();
                                            }
                                            else
                                            {
                                                cas.Add(colNames[1].ToString(),new SortedSet<string>());
                                                if (
                                                    !Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "CA Drawings")))
                                                    //make sure base folder is created
                                                {
                                                    Directory.CreateDirectory(Path.Combine(scoeBaseFolder,
                                                        secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "CA Drawings"));
                                                }
                                                currentCANumber = colNames[1];
                                                currentCA = colNames[2].Truncate(53).Trim();
                                                if (
                                                    !Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "CA Drawings",
                                                        currentCA.Replace(":", "-").Replace("/", "-").Replace("\"", "-") +
                                                        "_Empty"))) //create ca folder
                                                    //This also checks to make sure that the CA number.txt file exists in the folder
                                                {

                                                    Directory.CreateDirectory(Path.Combine(scoeBaseFolder,
                                                        secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "CA Drawings",
                                                        currentCA.Replace(":", "-").Replace("/", "-").Replace("\"", "-") +
                                                        "_Empty"));

                                                    System.IO.File.WriteAllText(
                                                        Path.Combine(scoeBaseFolder, secondaryProponent,
                                                            masterPlanningCategory,
                                                            fnameLong, "CA Drawings",
                                                            currentCA.Replace("/", "-")
                                                                .Replace(":", "-")
                                                                .Replace("\"", "-") +
                                                            "_Empty",
                                                            "CA Number - " +
                                                            colNames[1].Replace("\"", "-")
                                                                .Replace("/", "-")
                                                                .Replace(":", "-") +
                                                            ".txt"),
                                                        "Facility:\t" + colNames[1] + " " + Environment.NewLine +
                                                        "Description:\t" + colNames[2]);

                                                }
                                                else
                                                {
                                                    //just shows what the below evaluated true/false value will be
                                                    bool isCAtxtPresent =
                                                        Directory.EnumerateFiles(
                                                            Path.Combine(scoeBaseFolder, secondaryProponent,
                                                                masterPlanningCategory, fnameLong, "CA Drawings",
                                                                currentCA.Replace(":", "-")
                                                                    .Replace("/", "-")
                                                                    .Replace("\"", "-") +
                                                                "_Empty"), "*" + currentCANumber + "*",
                                                            SearchOption.AllDirectories).Any();

                                                    if (
                                                        !Directory.EnumerateFiles(
                                                            Path.Combine(scoeBaseFolder, secondaryProponent,
                                                                masterPlanningCategory, fnameLong, "CA Drawings",
                                                                currentCA.Replace(":", "-")
                                                                    .Replace("/", "-")
                                                                    .Replace("\"", "-") +
                                                                "_Empty"), "*" + currentCANumber + "*",
                                                            SearchOption.AllDirectories).Any())
                                                    {
                                                        int count = Directory.EnumerateFiles(
                                                            Path.Combine(scoeBaseFolder, secondaryProponent,
                                                                masterPlanningCategory, fnameLong, "CA Drawings",
                                                                currentCA.Replace(":", "-")
                                                                    .Replace("/", "-")
                                                                    .Replace("\"", "-") +
                                                                "_Empty"), "*" + currentCANumber + "*",
                                                            SearchOption.AllDirectories).Count();

                                                        Directory.CreateDirectory(Path.Combine(scoeBaseFolder,
                                                            secondaryProponent,
                                                            masterPlanningCategory, fnameLong, "CA Drawings",
                                                            currentCA.Replace(":", "-")
                                                                .Replace("/", "-")
                                                                .Replace("\"", "-") +
                                                            "_Empty") + (count + 1).ToString());

                                                        System.IO.File.WriteAllText(
                                                            Path.Combine(scoeBaseFolder, secondaryProponent,
                                                                masterPlanningCategory,
                                                                fnameLong, "CA Drawings",
                                                                currentCA.Replace("/", "-")
                                                                    .Replace(":", "-")
                                                                    .Replace("\"", "-") +
                                                                "_Empty" + (count + 1).ToString(),
                                                                "CA Number - " +
                                                                colNames[1].Replace("\"", "-")
                                                                    .Replace("/", "-")
                                                                    .Replace(":", "-") +
                                                                ".txt"),
                                                            "Facility:\t" + colNames[1] + " " + Environment.NewLine +
                                                            "Description:\t" + colNames[2]);

                                                    }
                                                }
                                            }
                                        }
                                    } while (reader.NextResult());
                                    reader.Close();
                                }
                            }
                            //FACILITY AND COMPONENTS
                            sql = string.Format(
                                @"SELECT DISTINCT File_Link, File_Nbr,File_Title from JCMS_File WHERE File_Id in 
                (SELECT File_Id FROM JCMS_File_Owner WHERE File_Owner_Id in  
                (SELECT Element_Id FROM Element WHERE Element_Nbr = '{0}')) 
                AND File_Class = 'DRAWING' AND File_Type = 'PDF' ORDER BY File_Nbr ASC;
                ", fname);
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                            {
                                SqlDataReader reader = cmd.ExecuteReader();
                                do
                                {
                                    while (reader.Read())
                                    {
                                        string[] colNames = new string[4];
                                        reader.GetValues(colNames);
                                        if (colNames[3] == null) //always true, left as is since I'm editing code
                                        {
                                            string drawPath = colNames[0];
                                            string drawName = colNames[1] + " - " +
                                                              colNames[2].Replace("\"", "-")
                                                                  .Replace("/", "-")
                                                                  .Replace(":", "-");
                                            drawName = drawName.Truncate(120);
                                            facs.Add(drawName);
                                            string shortDrawName = colNames[1];
                                            string finalDrawPath = drawDir.FullName + "\\" + drawName + ".pdf";
                                            try
                                            {
                                                if (drawPath ==
                                                    "JCMS\\DATA_FILES\\DRAWINGS\\PDF\\70229-4108M-MP101X.PDF")
                                                {
                                                    drawPath =
                                                        "JCMS\\DATA_FILES\\DRAWINGS\\PDF\\70229-4108_M-MP101X.PDF";
                                                    //fix, wrong in the database
                                                }
                                                File.Copy("C:" + "\\" + drawPath, finalDrawPath, true);
                                            }
                                            catch (Exception excpt)
                                            {

                                                Console.WriteLine("Error with file..."); //70229-4108_M-MP101X
                                                Console.WriteLine(drawName);
                                                Console.WriteLine();
                                            }
                                            //Now add a shortcut in that file
                                            string finalLinkPath = null;
                                            if (lookupToType.Contains("Facility") || lookupToType.Contains("facility"))
                                            {
                                                facs.Add(drawName);
                                                if (!Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                                    masterPlanningCategory, fnameLong, "Facility Drawings")))
                                                {
                                                    Directory.CreateDirectory(Path.Combine(scoeBaseFolder,
                                                        secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "Facility Drawings"));
                                                    System.IO.File.WriteAllText(
                                                        Path.Combine(scoeBaseFolder, secondaryProponent,
                                                            masterPlanningCategory, fnameLong, "Facility Drawings",
                                                            "Facility Description - " +
                                                            colNames[2].Replace("\"", "-")
                                                                .Replace("/", "-")
                                                                .Replace(":", "-") +
                                                            ".txt"),
                                                        "Facility:\t" + colNames[1] + " " + Environment.NewLine +
                                                        "Description:\t" + colNames[2]);
                                                }
                                                finalLinkPath = Path.Combine(scoeBaseFolder, secondaryProponent,
                                                    masterPlanningCategory, fnameLong, "Facility Drawings",
                                                    shortDrawName + ".pdf.lnk");
                                            }
                                            else if (lookupToType.Contains("Component"))
                                            {
                                                if (!Directory.Exists(Path.Combine(scoeBaseFolder, secondaryProponent,
                                                    masterPlanningCategory, fnameLong, "Component Drawings")))
                                                {
                                                    Directory.CreateDirectory(Path.Combine(scoeBaseFolder,
                                                        secondaryProponent,
                                                        masterPlanningCategory, fnameLong, "Component Drawings"));
                                                    System.IO.File.WriteAllText(
                                                        Path.Combine(scoeBaseFolder, secondaryProponent,
                                                            masterPlanningCategory, fnameLong, "Component Drawings",
                                                            "Component Number - " +
                                                            colNames[1].Replace("\"", "-")
                                                                .Replace("/", "-")
                                                                .Replace(":", "-") +
                                                            ".txt"),
                                                        "Facility:\t" + colNames[1] + " " + Environment.NewLine +
                                                        "Description:\t" + colNames[2]);
                                                }
                                                finalLinkPath = Path.Combine(scoeBaseFolder, secondaryProponent,
                                                    masterPlanningCategory, fnameLong, "Component Drawings",
                                                    shortDrawName + ".pdf.lnk");
                                            }
                                            WshShell wsh = new WshShell();
                                            IWshRuntimeLibrary.IWshShortcut shortcut =
                                                (IWshShortcut)
                                                    wsh.CreateShortcut(finalLinkPath);
                                            shortcut.Arguments = "";
                                            shortcut.TargetPath = finalDrawPath;
                                            // not sure about what this is for
                                            shortcut.WindowStyle = 1;
                                            if (lookupToType.Contains("Facility"))
                                            {
                                                shortcut.Description = "Facility Drawing";
                                            }
                                            else if (lookupToType.Contains("Component"))
                                            {
                                                shortcut.Description = "Component Drawing";
                                            }
                                            shortcut.WorkingDirectory = drawDir.FullName;
                                            shortcut.IconLocation = "icon location";
                                            shortcut.Save();
                                        }
                                        else
                                        {
                                            currentCA = colNames[2].Truncate(53).Trim();
                                        }
                                    }
                                } while (reader.NextResult());
                                reader.Close();
                            }
                            if (!Directory.EnumerateFileSystemEntries(fPath).Any())
                            {
                                try
                                {
                                    System.IO.Directory.Move(fPath, fPath + "_Empty");
                                }
                                catch (Exception excpt)
                                {
                                    Console.WriteLine("Failed to change folder name:" + fPath);
                                }
                            }
                            string htmlFile = "";
                            htmlFile = String.Format(@"
<!DOCTYPE html>
<html>
  <head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1"">
    <title>Facility Vetting</title>
    <link rel=""stylesheet"" href=""https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css"">
    <link rel=""stylesheet"" href=""https://cdnjs.cloudflare.com/ajax/libs/bulma/0.5.0/css/bulma.min.css"">
  </head>
  <body>
  <section class=""section"">
<section class=""hero is-medium is-primary is-bold"">
  <div class=""hero-body"">
    <div class=""container"">
      <h1 class=""title"">
        SCoE Vetting
      </h1>
      <h2 class=""subtitle"">
        Facility Number: {0}
      </h2>
    </div>
  </div>
</section>
      
<nav class=""breadcrumb has-arrow-separator"" aria-label=""breadcrumbs"">
  <ul>
    <li><a href=""#"">{1}</a></li>
    <li><a href=""#"">{2}</a></li>
    <li class=""is-active""><a href=""#"" aria-current=""page"">Facility</a></li>
  </ul>
</nav>
      
<div class=""card"">
  <header class=""card-header"">
    <p class=""card-header-title"">
      Description
    </p>
    <a class=""card-header-icon"">
      <span class=""icon"">
        <i class=""fa fa-angle-down""></i>
      </span>
    </a>
  </header>
  <div class=""card-content"">
    <div class=""content"">
      {3}
      <br>
      <small>Original Vetting Date: {4}</small>
    </div>
  </div>
  <footer class=""card-footer"">

  </footer>
</div>
<div style=""margin:20px;"">
</div>
      
    
    
      
      
      
      
      <div class=""tile is-ancestor"">
  <div class=""tile is-vertical is-8"">
    <div class=""tile"">
      <div class=""tile is-parent is-vertical"">
        <article class=""tile is-child notification is-primary"">
          <p class=""title"">Lookup to Noun</p>

<article class=""message is-primary is-medium"">
  <div class=""message-body"">
    {5}
  </div>
</article>


            
        </article>
        <article class=""tile is-child notification is-warning"">
          <p class=""title"">Lookup to Standard</p>
            
            
<article class=""message is-warning is-medium"">
  <div class=""message-body"">
    {8}
  </div>
</article>
    
    
        </article>
      </div>
      <div class=""tile is-parent"">
        <article class=""tile is-child notification is-info"">
          <p class=""title"">Lookup to Master Planning Category</p>

            <article class=""message is-primary is-medium"">
  <div class=""message-body"">
    {9}
  </div>
</article>
            
            
        </article>
      </div>
    </div>
    <div class=""tile is-parent"">
      <article class=""tile is-child notification is-danger"">
        <p class=""title"">Detail Field</p>

          
          <article class=""message is-danger is-medium"">
  <div class=""message-body"">
    {7}
  </div>
</article>
          
          
        <div class=""content"">
          <!-- Content -->
        </div>
      </article>
    </div>
  </div>
  <div class=""tile is-parent"">
    <article class=""tile is-child notification is-success"">
      <div class=""content"">
        <p class=""title"">Primary Construction Material</p>

          <article class=""message is-success is-medium"">
  <div class=""message-body"">
    {6}
  </div>
</article>
          
      </div>
    </article>
  </div>
</div>
      
      
      
      
      
      
      <article class=""message is-dark is-large"">
  <div class=""message-header"">
    <p>Proponent Comments</p>
  </div>
  <div class=""message-body"">
    {10}
  </div>
</article>


<article class=""message is-warning is-large"">
  <div class=""message-header"">
    <p>Proponent Recommendation</p>
  </div>
  <div class=""message-body"">
    {11}
  </div>
</article>
      
 
  </section>
      
      
      
      <article class=""media"">
  <figure class=""media-left"">
    <p class=""image is-64x64"">
    </p>
  </figure>
  <div class=""media-content"">
    <div class=""field"">
      <p class=""control"">
        <textarea rows=""20"" class=""textarea"" placeholder=""Design Agent Comments...""></textarea>
      </p>
    </div>
    <nav class=""level"">
      <div class=""level-left"">
        <div class=""level-item"">
        </div>
      </div>
      <div class=""level-right"">
        <div class=""level-item"">
          <label class=""checkbox"">
          </label>
        </div>
      </div>
    </nav>
  </div>
</article>
      
      
      
      
      
{12}
      
{13}

      <style>
      @media print  
{{
        div{{
        page-break-inside: avoid;
    }}
        article {{
        page-break-inside: avoid;
    }}
        footer{{
        page-break-inside: avoid;
    }}
}}
</style>
      
      
  </body>
    
    
    
    <footer class=""footer"">
  <div class=""container"">
    <div class=""content has-text-centered"">
      <p>
        An <strong>HDR</strong> document. 
      </p>
      <p>
      </p>
    </div>
  </div>
</footer>
</html>
        
                ", facilityNumber, textInfo.ToTitleCase(primaryProponent), textInfo.ToTitleCase(secondaryProponent), description, vettingDate, lookupToNoun, primaryConstructionMaterial, detailField, lookupToStandard, lookupToMasterPlanningCategory, newProponentComments,
                 proponentRecommendation, GetFacilityPdfString(facs), GetCaPdfString(cas));
                            try
                            {
                                File.WriteAllText(
                                    Path.Combine(scoeBaseFolder, secondaryProponent, masterPlanningCategory, fnameLong,
                                        "Cover Sheet.html"), htmlFile);
                            }
                            catch (Exception excpt)
                            {
                                File.WriteAllText(
                                    Path.Combine(scoeBaseFolder, secondaryProponent, masterPlanningCategory,
                                        fPath + "_Empty",
                                        "Cover Sheet.html"), htmlFile);
                            }
                        }
                    }
                }
                //Start of second program, 'Add Par to Folder'
                scoeBaseFolder = "C:\\SCoE";
                DirectoryInfo baseDi = new DirectoryInfo(scoeBaseFolder);
                foreach (var d in baseDi.GetDirectories("*"))
                {
                    if (d.Name != "Drawings")
                    {
                        foreach (var d1 in d.GetDirectories("*"))
                        {
                            foreach (var d2 in d1.GetDirectories("*"))
                            {
                                string[] types = {"A", "E", "G", "K", "M", "P", "S"};
                                string toAppend = " (";
                                foreach (string s in types)
                                {
                                    if (
                                        Directory.EnumerateFiles(d2.FullName, "*_" + s + "*",
                                            SearchOption.AllDirectories)
                                            .Any())
                                    {
                                        toAppend += s;
                                    }
                                }
                                toAppend += ")";
                                if (!d2.Name.ToString().Contains("_Empty"))
                                {
                                    //We might have a facility/comp with no files, but I never labeled comp or facilities as empty above
                                    if (
                                        !Directory.EnumerateFiles(d2.FullName, "*.pdf.lnk", SearchOption.AllDirectories)
                                            .Any())
                                        //if there are no pdfs, then label this file as empty
                                    {
                                        Directory.Move(d2.FullName, d2.FullName + "_Empty");
                                    }
                                    else
                                    {
                                        Directory.Move(d2.FullName, d2.FullName + toAppend);
                                    }
                                    var d2_changed = d1.GetDirectories("*" + d2.Name + "*")[0];
                                    if ((!d2_changed.ToString().Contains("()")) &&
                                        (!d2_changed.Name.ToString().Contains("_Empty")))
                                        //Don't let () or _Empty through this loop
                                    {
                                        foreach (var d3 in d2_changed.GetDirectories("*"))
                                        {
                                            if (d3.Name != "Component Drawings")
                                                //skip comp draw, since this is the only folder with those files
                                            {
                                                var v1 =
                                                    Directory.EnumerateDirectories(d3.Parent.FullName,
                                                        "*" + "CA Drawings" + "*",
                                                        SearchOption.AllDirectories).Any();
                                                var v2 =
                                                    Directory.EnumerateDirectories(d3.Parent.FullName,
                                                        "*" + "Facility Drawings" + "*",
                                                        SearchOption.AllDirectories).Any();
                                                var d3_changed = d3;
                                                if (
                                                    Directory.EnumerateDirectories(d3.Parent.FullName,
                                                        "*" + "CA Drawings" + "*",
                                                        SearchOption.AllDirectories).Any() &&
                                                    Directory.EnumerateDirectories(d3.Parent.FullName,
                                                        "*" + "Facility Drawings" + "*",
                                                        SearchOption.AllDirectories).Any())
                                                    //need both CA and Fac, otherwise do nothing, since the files are in one of the folders
                                                {
                                                    toAppend = " (";
                                                    foreach (string s in types)
                                                    {
                                                        if (
                                                            Directory.EnumerateFiles(d3.FullName, "*_" + s + "*",
                                                                SearchOption.AllDirectories).Any())
                                                        {
                                                            toAppend += s;
                                                        }
                                                    }
                                                    toAppend += ")";
                                                    Directory.Move(d3.FullName, d3.FullName + toAppend);
                                                    d3_changed = d2_changed.GetDirectories("*" + d3.Name + "*")[0];
                                                }
                                                if (!d3_changed.ToString().Contains("()"))
                                                {
                                                    int le = d3_changed.GetDirectories("CA Drawings").Length;
                                                    if (d3_changed.Name.Contains("CA Drawings") ||
                                                        d3_changed.GetDirectories("CA Drawings").Length > 0)
                                                    {
                                                        foreach (var d4 in d3_changed.GetDirectories("*"))
                                                        {
                                                            if (!d4.Name.Contains("_Empty"))
                                                            {
                                                                if (Directory.Exists(d4.FullName))
                                                                {
                                                                    var d4_changed = d3;
                                                                    toAppend = " (";
                                                                    foreach (string s in types)
                                                                    {

                                                                        if (
                                                                            Directory.EnumerateFiles(d4.FullName,
                                                                                "*_" + s + "*",
                                                                                SearchOption.AllDirectories).Any())
                                                                        {
                                                                            toAppend += s;
                                                                        }
                                                                    }
                                                                    toAppend += ")";
                                                                    Directory.Move(d4.FullName, d4.FullName + toAppend);
                                                                    d4_changed =
                                                                        d3_changed.GetDirectories("*" + d4.Name + "*")[0
                                                                            ];
                                                                    foreach (var d5 in d4_changed.GetDirectories("*"))
                                                                    {
                                                                        if (!d5.Name.Contains("_Empty"))
                                                                        {
                                                                            toAppend = " (";
                                                                            foreach (string s in types)
                                                                            {

                                                                                if (
                                                                                    Directory.EnumerateFiles(
                                                                                        d5.FullName,
                                                                                        "*_" + s + "*",
                                                                                        SearchOption.AllDirectories)
                                                                                        .Any())
                                                                                {
                                                                                    toAppend += s;
                                                                                }
                                                                            }
                                                                            toAppend += ")";
                                                                            Directory.Move(d5.FullName,
                                                                                d5.FullName + toAppend);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                foreach (var d in baseDi.GetDirectories("*",searchOption:SearchOption.AllDirectories))
                {
                    if (d.Name.Contains("CA Drawings"))
                    {
                        if (Directory.EnumerateFiles(d.FullName, "*.pdf*", SearchOption.AllDirectories).Any())
                        {

                        }
                        else
                        {
                            Directory.Move(d.FullName,d.FullName + "_Empty");
                        }
                    }
                }

            vWb.Close(false);
                xl.Quit();
            Console.WriteLine("Done!");
            }
            
        }
    }
}
