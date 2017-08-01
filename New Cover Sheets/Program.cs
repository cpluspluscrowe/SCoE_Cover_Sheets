using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using IWshRuntimeLibrary;
using Excel = Microsoft.Office.Interop.Excel;
using File = System.IO.File;
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
                                secondaryProponent = "NO_SECONDARY_PROPONENT";
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
                                masterPlanningCategory = "NO_MASTER_PLANNING_CATEGORY";
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
                            string newProponentComments = SpanInsert(vWs.Range["L" + i.ToString()].Value);


                            string designator = SpanInsert(vWs.Range["B" + i.ToString()].Value);//
                            string description = SpanInsert(vWs.Range["C" + i.ToString()].Value);//
                            string detailField = SpanInsert(vWs.Range["F" + i.ToString()].Value);//
                            string lookupToNoun = SpanInsert(vWs.Range["H" + i.ToString()].Value);//
                            string lookupToStandard = SpanInsert(vWs.Range["I" + i.ToString()].Value);//
                            string lookupToMasterPlanningCategory = SpanInsert(vWs.Range["J" + i.ToString()].Value);//
                            string primaryConstructionMaterial = SpanInsert(vWs.Range["E" + i.ToString()].Value);//
                            string primaryProponent = SpanInsert(vWs.Range["O" + i.ToString()].Value);//
                            string lookupToType = SpanInsert(vWs.Range["G" + i.ToString()].Value);//
                            string proponentRecommendation = SpanInsert(vWs.Range["K" + i.ToString()].Value);//

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
                                            if (lookupToType.Contains("Facility"))
                                            {
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
<head>
<title>JCMS Desktop</title>
<!-- Latest compiled and minified CSS -->
<link rel=""stylesheet"" href=""https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"" integrity=""sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u"" crossorigin=""anonymous"">

<!-- Optional theme -->
<link rel=""stylesheet"" href=""https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css"" integrity=""sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp"" crossorigin=""anonymous"">

<!-- Latest compiled and minified JavaScript -->
<script src=""https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"" integrity=""sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa"" crossorigin=""anonymous""></script>
</head>
<body id = ""main"">
<font face = ""arial"">
    <div class = ""wrapper"">
    <div class = ""LeftSide"">
    <div class = ""module"">
    <fieldset class=""fieldset-auto-width facilityNumber"">
      <legend>Facility Number</legend>
      <label>
       <span>{0}</span>
      </label>
    </fieldset>
    </div>
    <div class=""module2"">
    <fieldset class=""fieldset-auto-width"">
      <legend>Description</legend>
      <label>
        <span>{1}</span>
      </label>
    </fieldset>
    
    
    <fieldset class=""fieldset-auto-width"">
      <legend>Detail Field</legend>
      <label>
        <span>{2}</span>
      </label>
    </fieldset>
    </div>
        <div class = ""category"">
    <fieldset class=""fieldset-auto-width"">
      <legend>Lookup to Noun</legend>
      <label>
        <span>{3}</span>
      </label>
    </fieldset>
        
    <fieldset class=""fieldset-auto-width"">
      <legend>Lookup to Standard</legend>
      <label>
        <span>{4}</span>
      </label>
    </fieldset>
        
    <fieldset class=""fieldset-auto-width"">
      <legend>Lookup to Master Planning Category</legend>
      <label>
        <span>{5}</span>
      </label>
    </fieldset>
        
    <fieldset class=""fieldset-auto-width"">
      <legend>Primary Construction Material</legend>
      <label>
        <span>{6}</span>
      </label>
    </fieldset>
    </div>
    
        
        
</div>


    <div class = ""extra category2"">
            
        <div>
            <fieldset class = ""fieldset-auto-width"">
              <legend>Primary<br>Proponent</legend>
              <label>
                <span>{7}</span>
              </label>
            </fieldset>
<span class = ""arrow"">&#8594;</span>
            <fieldset class = ""fieldset-auto-width"">
              <legend>Secondary<br>Proponent</legend>
              <label>
                <span>{8}</span>
              </label>
            </fieldset>
<span class = ""arrow"">&#8594;</span>
            <fieldset class = ""fieldset-auto-width"">
              <legend>Lookup<br>to Type</legend>
              <label>
                <span>{9}</span>
              </label>
            </fieldset>
        </div>

    </div>


    <div class = ""Proponent description"" >
    <fieldset class=""long"">
      <legend>Proponent Comments</legend>
      <label>
        <span>{10}</span>
      </label>
    </fieldset>
        
    <fieldset class=""long"">
      <legend>Proponent Recommendation</legend>
      <label>
        <span>{11}</span>
      </label>
    </fieldset>
        
    </div>
    </div>
    </font>
</body>

<style>
    #main{{
        position:relative;
        overflow:hidden;
    }}
    #main h1, #main h3{{
        position:relative;
        z-index = 2;
    }}
    #main img{{
        position:absolute;
        width:100%;
        height:auto;
        opacity:0.07;
        background-size:cover;
    }}
    .facilityNumber{{
        border-color:rgba(106, 45, 38,1);
        color:rgba(106,45,38,1);
    }}
    fieldset{{
        padding:10px;
        margin:5px;
        border: 2px solid black;
        border-radius: 8px;
        height:auto;
    }}
    .fieldset-auto-width {{
         display: inline-block;
    }}
    .long{{
     height:auto;  
    width:auto;
    }}
.module{{
    background-color:rgba(208, 147, 140,.15);
    width:auto;
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
    }}
.module2{{
    background-color:rgba(200, 200, 100,.1);
    width:auto;
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
    }}
    .description{{
    width:auto;
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
    }}
div.fieldset-auto-width {{
    white-space: nowrap;
}}
    .Proponent{{
        float:right;
        top:0%;
        left:42%;
        width:55%;
        background-color:rgba(208, 147, 140,.15);
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
        margin:0;
        padding:0;
    }}
.FacNumber{{
        float:left;
        top:0%;
        background-color:rgba(208, 147, 140,.15);
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
        margin:0;
        padding:0;
    }}
    #table{{
    width:auto; 
    <!--background-color:rgba(208, 147, 140,.15);-->
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
    }}
    .folderStructure{{
        
    }}
    .category{{
    width:auto;
    background-color:rgba(150, 150, 191, .07);
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
    }}
    .LeftSide{{
        float:left;
        margin:0;
        padding:0;
        width:40%;
    }}
    #wrapper {{
     margin: 0 auto;
    position:relative;
}}
    legend{{
    font-size:24px;
    font-weight:600;
    }}
    label{{
     font-size:19px; 
    }}
.arrow{{
    font-size:200%;
    vertical-align:text-bottom;
    position:relative;
    top: -15px;
}}
.category2{{
    float:right;
    background-color:rgba(150, 150, 191, .07);
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
    }}
    .extra{{
        float:right;
        top:0%;
        left:42%;
        width:55%;
    webkit-border-radius: 15px;
    moz-border-radius: 10px; 
    border-radius: 7px;
        margin:0;
        padding:0;
    }}
</style>     
        
                ", facilityNumber, description, detailField, lookupToNoun, lookupToStandard,
                                lookupToMasterPlanningCategory, primaryConstructionMaterial, primaryProponent,
                                secondaryProponent, lookupToType, newProponentComments, proponentRecommendation);
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


            Console.WriteLine("Done!");
            }
        }
    }
}
