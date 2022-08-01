using ConsoleApp.OpenXML;
using System.Diagnostics;

using DocXToPdfConverter.DocXToPdfHandlers;

string today = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
string sourceFile = System.IO.Path.Combine("E:\\freeoutsource\\WebDevN-F\\ConsoleApp.OpenXML\\temps\\Maugiay-chung-nhan-dang-ky-ket-hon.docx");
string destinationFile = System.IO.Path.Combine("E:\\freeoutsource\\WebDevN-F\\ConsoleApp.OpenXML\\outputs\\", today + "_" + System.IO.Path.GetFileName(sourceFile));
string outputpdfFile = System.IO.Path.Combine("E:\\freeoutsource\\WebDevN-F\\ConsoleApp.OpenXML\\outputs\\", today + "_output.pdf");

#region
try
{
    // Create a copy of the template file and open the copy 
    File.Copy(sourceFile, destinationFile, true);

    // create key value pair, key represents words to be replace and 
    //values represent values in document in place of keys.
    Dictionary<string, string> Textplaceholder = new Dictionary<string, string>();
    Textplaceholder.Add("sohopdong", "07/22/1");
    Textplaceholder.Add("hotennguoia", "Họ tên Vợ");
    Textplaceholder.Add("hotenchong", "Họ tên Chồng");
    Textplaceholder.Add("ngaythangnama", "01-01-1998");
    Textplaceholder.Add("ngaythangnamb", "01-01-1991");
    Textplaceholder.Add("dantoca", "Kinh");
    Textplaceholder.Add("dantocb", "Kinh");
    Textplaceholder.Add("quocticha", "Việt Nam");
    Textplaceholder.Add("quoctichb", "Việt Nam");
    Textplaceholder.Add("noicutrua", "Hà Nội");
    Textplaceholder.Add("noicutrub", "Hà Nội");
    Textplaceholder.Add("giaytotuythana", "CMND số 12345678");
    Textplaceholder.Add("giaytotuythanchong", "CMND số 98765432");
    Textplaceholder.Add("noidangkykethon", "Ủy ban nhân dân phường Mộ Lao");
    Textplaceholder.Add("ngaythangdangkykethon", "22-07-2022");


    Wordprocessing.SearchAndReplace(destinationFile, Textplaceholder);

    string locationOfLibreOfficeSoffice = @"E:\freeoutsource\WebDevN-F\LibreOfficePortable\App\libreoffice\program\soffice.exe";

    LibreOfficeWrapper.Convert(destinationFile, outputpdfFile, locationOfLibreOfficeSoffice);


    ////Process.Start(destinationFile);
    //Process process = new Process();
    //// Configure the process using the StartInfo properties.
    //process.StartInfo.FileName = destinationFile;
    //process.StartInfo.Arguments = "-n";
    //process.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
    //process.Start();
    //process.WaitForExit();// Waits here for the process to exit.
}
catch (Exception ex)
{
    throw ex;
}

#endregion

Console.WriteLine("Hello, World!");
Console.ReadLine();

