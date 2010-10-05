using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net;
using System.IO;
using CookComputing.XmlRpc;
using System.Configuration;

namespace VrokSub
{
    class Program
    {
        private static IOSDb _osdbProxy;
        private static string CDFormat = ConfigurationManager.AppSettings["CDFormat"];
        private static string FileFormat = ConfigurationManager.AppSettings["FileFormat"];
        private static string FolderFormat = ConfigurationManager.AppSettings["FolderFormat"];
        private static string _theToken;

        private static List<string> movieFormats = new List<string>(ConfigurationManager.AppSettings["MovieFormat"].Split(','));
        private static List<string> subtitleFormats = new List<string>(ConfigurationManager.AppSettings["FolderFormat"].Split(','));
        private static string[] _langs;
        private static List<langMap> theLangMap = Utils.GenerateLangMap();
        private static int _dlCount;
        private static List<MovieFile> myFiles = new List<MovieFile>();

        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("");
                Console.WriteLine("vroksub.exe \"[Folder Path]\" [Language Code Sequence] [Params]");
                Console.WriteLine("");
                Console.WriteLine("Folder Path: The path to the folder that has all your movies. VrokSub will search all subfolders of this path for movies.");
                Console.WriteLine("");
                Console.WriteLine("Language Code Sequence: A sequence of two letter language codes according to your preference separated by coma. You can find the two letter codes here: http://www.loc.gov/standards/iso639-2/php/code_list.php. VrokSub will search subtitles of the first language code and if it doesn't find one it will continue to the next code.");
                Console.WriteLine("");
                Console.WriteLine("Params:");
                Console.WriteLine("");
                Console.WriteLine(" /rename will rename all the movies for which vroksub has found a subtitle using the format found in vroksub.exe.config");
                Console.WriteLine(" /newonly will only try to locate subtitles for movies without subtitles and ignore the ones that have subtitles");
                Console.WriteLine(" /nfo will download data from imdb.com (like actors, directors etc) and save them to a nfo file named like your movie");
                Console.WriteLine(" /covers will download the cover images imdb uses and save them to a jpg file named like your movie");
                Console.WriteLine(" /folders will create a folder for each movie and move all files there. If /covers is used with this then a folder.jpg will also be created");
                Console.WriteLine(" /nosubfolders will not search every subfolder of the given folder for movie files");
                Console.WriteLine(" /move=\"[Output Path]\" will move all files and folders to the given path. Useful if combined with folders");
                Console.WriteLine(" /nolanguageinfilename will not add the language to the filename. Requires that you only ask for subtitles in one language");
                Console.WriteLine("");
                Console.WriteLine("Examples:");
                Console.WriteLine("vroksub.exe \"c:\\my videos\" gr,en");
                Console.WriteLine(" This will first try to locate a greek subtitle for every movie in c:\\my videos (including subfolders) and if it doesn't find one it will try to find one in english. You can use more language codes if you'd like.");
                Console.WriteLine("");
                Console.WriteLine("vroksub.exe \"c:\\my videos\" it /rename");
                Console.WriteLine(" will also rename all movies where subtitle has been found with year and cd number: MovieName(Year)-CD2.avi");
                Console.WriteLine("");
                Console.WriteLine("vroksub.exe \"c:\\my videos\" de /newonly");
                Console.WriteLine(" will only search for subtitles for movies that don't have one already.");
                Console.WriteLine("");
                Console.WriteLine("vroksub.exe \"c:\\unsubbedvideos\" nl /rename /folders /nfo /covers /move=\"c:\\my videos\"");
                Console.WriteLine(" 1) Creates a folder for each movie under c:\\my videos with the name format found in vroksub.exe.config");
                Console.WriteLine(" 2) Renames each movie using the format found in vroksub.exe.config");
                Console.WriteLine(" 3) Downloads dutch subtitles for each movie");
                Console.WriteLine(" 4) Downloads imdb details and saves them to the output folder");
                Console.WriteLine(" 5) Downloads imdb covers and saves them to the output folder as folder.jpg and movie.jpg");
                Console.WriteLine(" Note: Only movies with found subtitles will be affected");

                Console.ReadLine();
            }
            else
            {
                List<string> argList = new List<string>(args);
                argList.ForEach(ToLower);
                bool ren = false;
                bool newOnly = false;
                bool nfo = false;
                bool folders = false;
                bool imdb = false;
                bool covers = false;
                bool noLanguageInFilename = false;
                string inputPath;
                string outputPath = "";
                string langSeq;
                foreach (string arg in argList)
                {
                    if (arg.StartsWith("/move="))
                    {
                        outputPath = arg.Replace("/move=", "");
                    }
                }
                if (args.Length > 2)
                {
                    ren = argList.Contains("/rename");
                    newOnly = argList.Contains("/newonly");
                    nfo = argList.Contains("/nfo");
                    folders = argList.Contains("/folders");
                    covers = argList.Contains("/covers");
                    imdb = (nfo || covers);
                    noLanguageInFilename = argList.Contains("/nolanguageinfilename");
                }
                argList.Remove("/nolanguageinfilename");
                argList.Remove("/rename");
                argList.Remove("/newonly");
                argList.Remove("/nfo");
                argList.Remove("/folders");
                argList.Remove("/covers");
                argList.Remove("/move=" + outputPath);
                if (Directory.Exists(args[1]) && (!Directory.Exists(args[0])))
                {
                    inputPath = args[1];
                    langSeq = args[0];
                }
                else
                {
                    inputPath = args[0];
                    langSeq = args[1];
                }

                if (langSeq.Contains(",") && noLanguageInFilename)
                {
                    Console.WriteLine("Multiple countrycodes not supported with parameter /nolanguageinfilename");
                    Console.ReadLine();
                    return;
                }

                Go(inputPath, outputPath, Get2CodeStr(langSeq), false, ren, newOnly, nfo, folders, imdb, covers, noLanguageInFilename);
            }
        }

        private static void ToLower(string a)
        {
            a = a.ToLower();
        }

        private static string Get3CodeStr(string the2CodeStr)
        {
            string c3 = "";
            foreach (string l2Code in the2CodeStr.Split(','))
            {
                foreach (langMap l in theLangMap)
                {
                    if (l.two.Contains(l2Code))
                    {
                        if (!c3.Contains(l.three))
                        {
                            c3 += "," + l.three;
                        }
                    }
                }
            }
            if (c3 != "") { c3 = c3.Substring(1); }
            return c3;
        }

        private static string Get2CodeStr(string the2CodeStr)
        {
            string c3 = "";
            foreach (string l2Code in the2CodeStr.Split(','))
            {
                foreach (langMap l in theLangMap)
                {
                    if (l.two.Contains(l2Code))
                    {
                        if (!c3.Contains(l.two))
                        {
                            c3 += "," + l.two;
                        }
                    }
                }
            }
            if (c3 != "") { c3 = c3.Substring(1); }
            return c3;
        }

        private static void Go(string folderArg, string outputPath, string langArg, bool overwrite, bool rename, bool newOnly, bool nfo, bool folders, bool imdb, bool covers, bool noLanguageInFilename)
        {
            Console.WriteLine("Searching for movie files");
            #region Get Movie Files
            if (File.Exists(folderArg))
            {
                if (outputPath == "")
                {
                    outputPath = Path.GetDirectoryName(folderArg);
                }
                MovieFile theFile = new MovieFile();
                theFile.filename = folderArg;
                theFile.getOldSubtitle(subtitleFormats, theLangMap, folderArg);
                myFiles.Add(theFile);
            }
            else if (Directory.Exists(folderArg))
            {
                if (outputPath == "")
                {
                    outputPath = folderArg;
                }
                foreach (string extension in movieFormats)
                {
                    foreach (string filename in Utils.GetFilesByExtensions(folderArg, "." + extension, SearchOption.AllDirectories))
                    {
                        MovieFile theFile = new MovieFile();
                        theFile.filename = filename;
                        theFile.getOldSubtitle(subtitleFormats, theLangMap, folderArg);
                        myFiles.Add(theFile);
                    }
                }
                Console.WriteLine("Found " + myFiles.Count + " Movie Files");
            }
            else
            {
                throw new Exception("The folder or file given does not exist");
            }
            #endregion

            Console.WriteLine("Creating subtitle request from movie files");
            #region Create subtitle request from movie files
            _langs = Get3CodeStr(langArg).Split(',');
            List<string> l3 = new List<string>(_langs);
            subInfo[] si = new subInfo[myFiles.Count * _langs.Length];
            int i = 0;
            foreach (MovieFile theFile in myFiles)
            {
                if ((newOnly && (theFile.oldSubtitle == "")) || (!newOnly))
                {
                    foreach (string lang in _langs)
                    {
                        try
                        {
                            si[i] = new subInfo();
                            si[i].sublanguageid = lang;
                            theFile.hash = Utils.ToHexadecimal(Utils.ComputeMovieHash(theFile.filename));
                            si[i].moviehash = theFile.hash;
                            si[i].moviebytesize = new FileInfo(theFile.filename).Length.ToString();
                            i++;

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error with file: " + theFile.filename + "\n" + ex.Message);
                        }
                    }
                }
            }
            #endregion

            Console.WriteLine("Connecting...");
            #region Connect and login
            _osdbProxy = XmlRpcProxyGen.Create<IOSDb>();
            _osdbProxy.Url = "http://api.opensubtitles.org/xml-rpc";
            _osdbProxy.KeepAlive = false;
            if (myFiles.Count == 0) { Console.WriteLine("No movies found"); }
            XmlRpcStruct login = _osdbProxy.LogIn("", "", "en", "vroksub");
            _theToken = login["token"].ToString();
            #endregion

            Console.WriteLine("Searching for subtitles...");
            #region Search for subtitles
            subrt subResults = null;
            try
            {
                subResults = _osdbProxy.SearchSubtitles(_theToken, si);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error when searching for subtitle: " + e.Message);
            }
            #endregion

            Console.WriteLine("Found subtitles:");
            #region Choose best subtitle
            if (subResults != null)
            {
                BindingList<subRes> g = new BindingList<subRes>(subResults.data);
                _dlCount = 0;
                foreach (MovieFile mf in myFiles)
                {
                    try
                    {
                        if ((newOnly && (mf.oldSubtitle == "")) || (!newOnly))
                        {
                            if (mf.SelectBestSubtitle(g, l3))
                            {
                                Console.WriteLine(mf.subRes.MovieName + " - " + mf.subRes.LanguageName);
                                _dlCount++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Could not choose subtitle for: " + mf.filename + "\n" + ex.Message);
                    }
                }
            #endregion

                Console.WriteLine("Downloading subtitles...");
                #region Download Subtitles
                string[] ids = new string[_dlCount];
                int k = 0;
                foreach (MovieFile myf in myFiles)
                {
                    if (myf.subtitleId != null)
                    {
                        ids[k] = myf.subtitleId;
                        k++;
                    }
                }
                subdata files = _osdbProxy.DownloadSubtitles(_theToken, ids);
                #endregion

                if (imdb)
                {
                    Console.WriteLine("Fetching imdb details...");
                    #region Fetch imdb details
                    foreach (MovieFile myf in myFiles)
                    {
                        if (myf.subRes != null)
                        {
                            try
                            {
                                myf.imdbinfo = _osdbProxy.GetIMDBMovieDetails(_theToken, "0" + myf.subRes.IDMovieImdb).data;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error fetching imdb data for: " + myf.filename + "\n" + ex.Message);
                            }
                        }
                    }
                    #endregion
                }

                Console.WriteLine("Saving subtitles...");

                #region Process (rename, create folders, savenfo and save) subtitles
                foreach (subtitle s in files.data)
                {
                    foreach (MovieFile m in myFiles)
                    {
                        try
                        {
                            if (m.subtitleId == s.idsubtitlefile)
                            {
                                m.subtitle = Utils.DecodeAndDecompress(s.data);
                                if (outputPath != folderArg)
                                {
                                    if (!Directory.Exists(outputPath))
                                    {
                                        Directory.CreateDirectory(outputPath);
                                    }
                                    if (!File.Exists(outputPath + "\\" + Path.GetFileName(m.filename)))
                                    {
                                        try
                                        {
                                            File.Move(m.filename, outputPath + "\\" + Path.GetFileName(m.filename));
                                            m.originalfilename = Path.GetFileName(m.filename);
                                            m.filename = outputPath + "\\" + Path.GetFileName(m.filename);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("Error moving movie: " + m.filename + "\n" + ex.Message);
                                        }
                                    }
                                }
                                if (folders)
                                {
                                    m.newFolder(outputPath, FolderFormat);
                                }
                                if (rename)
                                {
                                    m.rename(FileFormat, CDFormat);
                                }
                                if (nfo)
                                {
                                    m.saveNfo();
                                }
                                m.SaveSubtitle(overwrite, noLanguageInFilename);
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error saving subtitle for: " + m.filename + "\n" + ex.Message);
                        }
                    }
                }
            }
                #endregion

            if (covers)
            {
                Console.WriteLine("Downloading covers...");
                #region Download covers
                WebClient client = new WebClient();

                foreach (MovieFile myf in myFiles)
                {
                    if (myf.imdbinfo != null)
                    {
                        if ((myf.imdbinfo.cover != null) && (myf.imdbinfo.cover != ""))
                        {
                            try
                            {
                                Stream strm = client.OpenRead(myf.imdbinfo.cover);
                                FileStream writecover = new FileStream(Path.GetDirectoryName(myf.filename) + "\\" + Path.GetFileNameWithoutExtension(myf.filename) + ".jpg", FileMode.Create);
                                int a;
                                do
                                {
                                    a = strm.ReadByte();
                                    writecover.WriteByte((byte)a);
                                }
                                while (a != -1);
                                writecover.Position = 0;
                                File.Copy(Path.GetDirectoryName(myf.filename) + "\\" + Path.GetFileNameWithoutExtension(myf.filename) + ".jpg", Path.GetDirectoryName(myf.filename) + "\\" + Path.GetFileName(myf.filename) + ".jpg");
                                if (folders)
                                {
                                    File.Copy(Path.GetDirectoryName(myf.filename) + "\\" + Path.GetFileNameWithoutExtension(myf.filename) + ".jpg", Path.GetDirectoryName(myf.filename) + "\\folder.jpg");
                                }
                                strm.Close();
                                writecover.Close();
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error saving cover for: " + myf.filename + "\n" + ex.Message);
                            }
                        }
                    }
                #endregion
                }
            }

            Console.WriteLine("Disconnecting...");
            #region Disconnect
            _osdbProxy.LogOut(_theToken);
            #endregion

        }

    }

}
