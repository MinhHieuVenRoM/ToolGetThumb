using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;

namespace ToolNews
{
    class Program
    {
        /// <summary>
        /// PathProject Link chứa output img
        /// URL link nhập vào.
        /// </summary>
        static string PathProject = "C:\\Users\\User\\source\\repos\\ToolNews";
        private static System.IO.FileStream stream;
        private static List<String> DSIdNewsError = new List<string>();
        private static List<String> DSlinkNewsError = new List<string>();


        public static string URL = "";
        static void Main(string[] args)
        {
            string opt;
            Console.WriteLine("Tool Get Thumb from News");

            Console.WriteLine("Input path, if has a file contains id of post error, please enter before," +
                              " otherwise, enter the link to the full file to be filtered ");


            URL = Console.ReadLine();


            Console.WriteLine(" 1 . Get link news not thumb ");
            Console.WriteLine(" 2 . Remder thumb from list error from list id of post");
            Console.WriteLine(" 3 . Remder thumb from list full data");

            opt = Console.ReadLine();

            if (opt.Equals("1"))
            {
                // get link and create excel file containing list error
                ImportExcel_quickly(URL);

            }
            if (opt.Equals("2"))
            {

                // ghi tu file danh sach loi và file full
                ImportExcel_scanLink(URL);

            }
            else
            {
                //ghi từ file full 
                ImportExcel_getLink(URL);
            }



        }
        /// <summary>
        /// Tạo ra file excel chưa id và link của các NewsID chưa có thumb
        /// </summary>
        /// <param name="FilePath">Path của danh sách đày đủ cần kiểm tra thumb</param>
        /// <returns></returns>
        public static int ImportExcel_quickly(String FilePath)
        {
            // Encodeing 1252 for .Netcore , Need Install CodePagesEncodingProvider

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            int rowexecl = 1;
            try
            {
                string filePath = "";
                if (FilePath != null && FilePath.Length > 0)
                {
                    filePath = FilePath;
                }
                else
                {
                    Console.WriteLine("File Path không đúng");
                    return 2;
                }

                stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;

                //1. Reading Excel file
                if (Path.GetExtension(filePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                //3. DataSet - Create column names from first row
                // excelReader.IsFirstRowAsColumnNames = true; 

                DataTable dataRecords = new DataTable();
                dataRecords = result.Tables[0];
                dataRecords.Rows[0].Delete();
                dataRecords.AcceptChanges();
                List<PostModel> reqData = new List<PostModel>();
                if (dataRecords.Rows.Count > 0)
                {

                    foreach (DataRow row in dataRecords.Rows)
                    {
                        if (row[0] != null && row[0].ToString().Trim().Length > 0)
                        {
                            PostModel exceldata = new PostModel();
                            rowexecl++;
                            exceldata.NEWSID = row[0].ToString();
                            exceldata.THUMBNAILIMAGE = row[1].ToString();
                            exceldata.DETAILIMAGE = row[2].ToString();
                            exceldata.TITLE = row[3].ToString();
                            exceldata.URL = row[4].ToString();
                            exceldata.METATITLE = row[5].ToString();
                            exceldata.METAKEYWORD = row[6].ToString();
                            exceldata.METADESCRIPTION = row[7].ToString();
                            exceldata.CREATEDDATE = Convert.ToDateTime(row[8]);
                            exceldata.CREATEDUSER = row[9].ToString();
                            exceldata.CREATEDCUSTOMERID = row[10].ToString();
                            exceldata.TAGS = row[11].ToString();
                            exceldata.LISTCATEGORYID = row[12].ToString();
                            exceldata.LISTCATEGORYNAME = row[13].ToString();
                            reqData.Add(exceldata);

                            String Url = "https://cdn.tgdd.vn/Files/";

                            string Year = exceldata.CREATEDDATE.Year.ToString();
                            string month = exceldata.CREATEDDATE.Month.ToString();
                            if (month.Length == 1)
                            {
                                month = "0" + month;

                            }
                            string day = exceldata.CREATEDDATE.Day.ToString();
                            if (day.Length == 1)
                            {
                                day = "0" + day;

                            }
                            Url += Year + "/" + month + "/" + day + "/" + exceldata.NEWSID + "/";
                            if (!exceldata.DETAILIMAGE.Equals("null") && !exceldata.DETAILIMAGE.Contains("https://"))
                            {
                                string arrListStr = exceldata.DETAILIMAGE.Replace(".jpg", "_300x300.jpg").Replace(".png", "_300x300.png").Replace(".jpeg", "_300x300.jpeg");
                                Url += arrListStr;
                                GetPage(Url, exceldata.NEWSID);

                            }
                            if (exceldata.DETAILIMAGE.Equals("null") && !exceldata.THUMBNAILIMAGE.Equals("null") && !exceldata.DETAILIMAGE.Contains("https://"))

                            {
                                string arrListStr = exceldata.THUMBNAILIMAGE.Replace(".jpg", "_300x300.jpg").Replace(".png", "_300x300.png").Replace(".jpeg", "_300x300.jpeg");
                                Url += arrListStr;
                                GetPage(Url, exceldata.NEWSID);
                            }




                        }
                    }

                }

                stream.Close();
                stream.Dispose();
                ExportExcel();




                return 1;

            }
            catch (Exception e)
            {
                stream.Close();
                stream.Dispose();
                Console.WriteLine("File bị lỗi, vui lòng kiểm tra lại file , chú ý dòng " + rowexecl);
                return 2;
            }
        }

        /// <summary>
        /// Tạo thumb từ danh sách đầy đủ 
        /// </summary>
        /// <param name="FilePath">  Path của danh sách đày đủ cần tạo thumb khi đã có DETAILIMAGE hoặc THUMBNAILIMAGE</param>
        /// <returns></returns>
        public static int ImportExcel_getLink(String FilePath)
        {
            // Encodeing 1252 for .Netcore , Need Install CodePagesEncodingProvider

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            int rowexecl = 1;
            try
            {
                string filePath = "";
                if (FilePath != null && FilePath.Length > 0)
                {
                    filePath = FilePath;
                }
                else
                {
                    Console.WriteLine("File Path không đúng");
                    return 2;
                }

                stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;

                //1. Reading Excel file
                if (Path.GetExtension(filePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                //3. DataSet - Create column names from first row
                // excelReader.IsFirstRowAsColumnNames = true; 

                DataTable dataRecords = new DataTable();
                dataRecords = result.Tables[0];
                dataRecords.Rows[0].Delete();
                dataRecords.AcceptChanges();
                List<PostModel> reqData = new List<PostModel>();
                if (dataRecords.Rows.Count > 0)
                {

                    foreach (DataRow row in dataRecords.Rows)
                    {
                        if (row[0] != null && row[0].ToString().Trim().Length > 0)
                        {
                            PostModel exceldata = new PostModel();
                            rowexecl++;
                            exceldata.NEWSID = row[0].ToString();
                            exceldata.THUMBNAILIMAGE = row[1].ToString();
                            exceldata.DETAILIMAGE = row[2].ToString();
                            exceldata.TITLE = row[3].ToString();
                            exceldata.URL = row[4].ToString();
                            exceldata.METATITLE = row[5].ToString();
                            exceldata.METAKEYWORD = row[6].ToString();
                            exceldata.METADESCRIPTION = row[7].ToString();
                            exceldata.CREATEDDATE = Convert.ToDateTime(row[8]);
                            exceldata.CREATEDUSER = row[9].ToString();
                            exceldata.CREATEDCUSTOMERID = row[10].ToString();
                            exceldata.TAGS = row[11].ToString();
                            exceldata.LISTCATEGORYID = row[12].ToString();
                            exceldata.LISTCATEGORYNAME = row[13].ToString();


                            String Url = "https://cdn.tgdd.vn/Files/";

                            string Year = exceldata.CREATEDDATE.Year.ToString();
                            string month = exceldata.CREATEDDATE.Month.ToString();
                            if (month.Length == 1)
                            {
                                month = "0" + month;

                            }
                            string day = exceldata.CREATEDDATE.Day.ToString();
                            if (day.Length == 1)
                            {
                                day = "0" + day;

                            }
                            Url += Year + "/" + month + "/" + day + "/" + exceldata.NEWSID + "/";
                            if (!exceldata.DETAILIMAGE.Equals("null") && !exceldata.DETAILIMAGE.Contains("https://"))
                            {

                                Url += exceldata.DETAILIMAGE;
                                string arrListStr = exceldata.DETAILIMAGE.Replace(".jpg", "_300x300.jpg").Replace(".png", "_300x300.png").Replace(".jpeg", "_300x300.jpeg");
                               var  namefile = arrListStr;


                                String pathimg = PathProject + "\\NewsFull\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID;
                                System.IO.DirectoryInfo dinew = new DirectoryInfo(pathimg);
                                dinew.Create();
                                GetFileFromUrl(pathimg + "\\" + namefile, Url);
                                Image image = Image.FromFile(pathimg + "\\" + namefile);

                                String pathimgNews = PathProject + "\\News\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID + "\\";
                                System.IO.DirectoryInfo di = new DirectoryInfo(pathimgNews);
                                di.Create();

                                ResizeAndCompress(image, pathimgNews + namefile, 300, 300);


                            }
                            if (exceldata.DETAILIMAGE.Equals("null") && !exceldata.THUMBNAILIMAGE.Equals("null") && !exceldata.DETAILIMAGE.Contains("https://"))

                            {


                                Url += exceldata.THUMBNAILIMAGE;
                                string arrListStr = exceldata.THUMBNAILIMAGE.Replace(".jpg", "_300x300.jpg").Replace(".png", "_300x300.png").Replace(".jpeg", "_300x300.jpeg");

                              
                                var namefile = arrListStr;

                                String pathimg = PathProject + "\\NewsFull\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID;
                                System.IO.DirectoryInfo dinew = new DirectoryInfo(pathimg);
                                dinew.Create();
                                GetFileFromUrl(pathimg + "\\" + namefile, Url);
                                Image image = Image.FromFile(pathimg + "\\" + namefile);

                                String pathimgNews = PathProject + "\\News\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID + "\\";
                                System.IO.DirectoryInfo di = new DirectoryInfo(pathimgNews);
                                di.Create();

                                ResizeAndCompress(image, pathimgNews + namefile, 300, 300);



                            }

                            exceldata.Link = Url;
                            reqData.Add(exceldata);

                        }
                    }

                }

                stream.Close();
                stream.Dispose();
                return 1;

            }
            catch (Exception e)
            {
                stream.Close();
                stream.Dispose();
                Console.WriteLine("File bị lỗi, vui lòng kiểm tra lại file , chú ý dòng " + rowexecl);
                return 2;
            }
        }
        /// <summary>
        /// Get link URL 200 OK hay lỗi.
        /// </summary>
        /// <param name="url"></param>
        /// <param name="idNews"></param>
        public static void GetPage(String url, string idNews)
        {
            try
            {
                // Creates an HttpWebRequest for the specified URL.
                System.Net.HttpWebRequest myHttpWebRequest = (System.Net.HttpWebRequest)WebRequest.Create(url);
                // Sends the HttpWebRequest and waits for a response.
                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                if (myHttpWebResponse.StatusCode == HttpStatusCode.OK)
                    Console.WriteLine("\r\nStatus Code is OK " + url);
                else
                {
                    DSIdNewsError.Add(idNews);
                }

                // Releases the resources of the response.
                myHttpWebResponse.Close();
            }
            catch (WebException e)
            {

                DSIdNewsError.Add(idNews);
                DSlinkNewsError.Add(url);
                Console.WriteLine("Error " + url);
            }
            catch (Exception e)
            {
                DSlinkNewsError.Add(url);
                DSIdNewsError.Add(idNews);
                Console.WriteLine("Error " + url);
            }
        }

        /// <summary>
        /// Ghi ra file excel
        /// </summary>
        public static void ExportExcel()
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //create a new Worksheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                worksheet.Cells[1, 1].Value = "NewsId";
                worksheet.Cells[1, 2].Value =  "Link render _300x300";
                //add some text to cell A1
                for (int i = 1; i < DSIdNewsError.Count; i++)
                {
                    worksheet.Cells[i + 1, 1].Value = DSIdNewsError[i];
                    worksheet.Cells[i + 1, 2].Value = DSlinkNewsError[i];

                }


                //the path of the file
                string filePath = PathProject + "\\FileThumbError\\ListSuccess.xlsx";

                //Write the file to the disk
                FileInfo fi = new FileInfo(filePath);
                excelPackage.SaveAs(fi);
                Console.WriteLine("Link File " + PathProject + "\\FileThumbError\\ListSuccess.xlsx");
            }

        }


        #region Image Extenstion
        public static string PickColorOfImg(string imgUrl)
        {
            try
            {
                WebRequest request = WebRequest.Create(imgUrl);
                WebResponse response = request.GetResponse();
                System.IO.Stream responseStream = response.GetResponseStream();
                if (responseStream != null)
                {
                    System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(responseStream);
                    var color = bitmap.GetPixel(0, 0);
                    return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
                }
                return string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        public static void ResizeAndCompress(Image srcImage, string outputFile, int newWidth, int newHeight)
        {
            try
            {
                using (var newImage = new Bitmap(newWidth, newHeight))
                using (var graphics = Graphics.FromImage(newImage))
                {
                    // Resize
                    graphics.SmoothingMode = SmoothingMode.HighSpeed;
                    graphics.InterpolationMode = InterpolationMode.Default;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighSpeed;
                    graphics.DrawImage(srcImage, new Rectangle(0, 0, newWidth, newHeight));

                    // Compress
                    ImageCodecInfo encoder;
                    if (outputFile.Contains(".png"))
                    {
                        encoder = GetEncoder(ImageFormat.Png);
                    }
                    else
                    {
                        encoder = GetEncoder(ImageFormat.Jpeg);
                    }

                    Encoder myEncoder = Encoder.Quality;
                    EncoderParameters myEncoderParameters = new EncoderParameters(1);
                    EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 90L);
                    myEncoderParameters.Param[0] = myEncoderParameter;

                    // Save
                    //newImage.Save(outputFile);
                    newImage.Save(outputFile, encoder, myEncoderParameters);
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;
#endif
            }
        }

        public static void ResizeAndCompress(string imageFile, string outputFile, int maxWidth, int maxHeight)
        {
            using (var srcImage = Image.FromFile(imageFile))
            {
                var newWidth = srcImage.Width > maxWidth ? maxWidth : srcImage.Width;
                var scaleFactor = Math.Round((decimal)newWidth / srcImage.Width, 2);
                var newHeight = (int)(srcImage.Height * scaleFactor);
                if (newHeight > maxHeight)
                {
                    newHeight = maxHeight;
                    scaleFactor = Math.Round((decimal)newHeight / srcImage.Height, 2);
                    newWidth = (int)(srcImage.Width * scaleFactor);
                }

                ResizeAndCompress(srcImage, outputFile, newWidth, newHeight);
            }
        }

        public static void ResizeAndCompress(string imageFile, string outputFile, double scaleFactor)
        {
            using (var srcImage = Image.FromFile(imageFile))
            {
                var newWidth = (int)(srcImage.Width * scaleFactor);
                var newHeight = (int)(srcImage.Height * scaleFactor);
                ResizeAndCompress(srcImage, outputFile, newWidth, newHeight);
            }
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        /// <summary>
        /// Get hình thumb để hiển thị trên mobile (300x300)
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string GetThumbMobile(string url)
        {
            if (string.IsNullOrEmpty(url))
                return string.Empty;

            return url.Replace(".jpg", "_300x300.jpg").Replace(".png", "_300x300.png").Replace(".jpeg", "_300x300.jpeg");
        }

        #endregion


        /// <summary>
        /// Lấy Ảnh từ URL
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="url"></param>
        static void GetFileFromUrl(string fileName, string url)
        {
            byte[] content;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            WebResponse response = request.GetResponse();

            Stream stream = response.GetResponseStream();
            using (BinaryReader reader = new BinaryReader(stream))
            {
                // kích thước tối đa get về đc 5mb
                content = reader.ReadBytes(5000000);
                reader.Close();
            }
            response.Close();
            FileStream fs = new FileStream(fileName, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            try
            {
                bw.Write(content);

            }
            finally
            {
                bw.Close();
                fs.Close();
            }
        }
        public static void SaveImage(string imageUrl, string filename, ImageFormat format)
        {
            WebClient client = new WebClient();
            Stream stream = client.OpenRead(imageUrl);
            Bitmap bitmap; bitmap = new Bitmap(stream);

            if (bitmap != null)
            {
                bitmap.Save(filename, format);
            }

            stream.Flush();
            stream.Close();
            client.Dispose();
        }

        ///<summary>
        /// Get Danh sách các ID lỗi vào một file data tổng
        ///</summary>
        public static int ImportExcel_scanLink(String FilePath)
        {
            // Encodeing 1252 for .Netcore , Need Install CodePagesEncodingProvider

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            int rowexecl = 1;
            try
            {
                string filePath = "";
                if (FilePath != null && FilePath.Length > 0)
                {
                    filePath = FilePath;
                }
                else
                {
                    Console.WriteLine("File Path không đúng");
                    return 2;
                }

                stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;

                //1. Reading Excel file
                if (Path.GetExtension(filePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                //3. DataSet - Create column names from first row
                // excelReader.IsFirstRowAsColumnNames = true; 

                DataTable dataRecords = new DataTable();
                dataRecords = result.Tables[0];
                dataRecords.Rows[0].Delete();
                dataRecords.AcceptChanges();
                List<PostModel> reqData = new List<PostModel>();
                if (dataRecords.Rows.Count > 0)
                {

                    foreach (DataRow row in dataRecords.Rows)
                    {
                        if (row[0] != null && row[0].ToString().Trim().Length > 0)
                        {
                            PostModel exceldata = new PostModel();
                            rowexecl++;
                            exceldata.NEWSID = row[0].ToString();
                            reqData.Add(exceldata);





                        }
                    }

                }

                Console.WriteLine("Input path full data");
                var linkfull = Console.ReadLine();
                Console.WriteLine("Runnnnnnnnnnnnnnn");


                //Map danh sách với Danh sách tổng
                scanLinkErrro(reqData, linkfull);


                stream.Close();
                stream.Dispose();

                // In ra danh sách đã làm được
                ExportExcel();




                return 1;

            }
            catch (Exception e)
            {
                stream.Close();
                stream.Dispose();
                Console.WriteLine("File bị lỗi, vui lòng kiểm tra lại file , chú ý dòng " + rowexecl);
                return 2;
            }
        }


        public static int scanLinkErrro(List<PostModel> list, string FilePath)
        {
            // Encodeing 1252 for .Netcore , Need Install CodePagesEncodingProvider

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            int rowexecl = 1;
            try
            {
                string filePath = "";
                if (FilePath != null && FilePath.Length > 0)
                {
                    filePath = FilePath;
                }
                else
                {
                    Console.WriteLine("File Path không đúng");
                    return 2;
                }

                stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;

                //1. Reading Excel file
                if (Path.GetExtension(filePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                //3. DataSet - Create column names from first row
                // excelReader.IsFirstRowAsColumnNames = true; 

                DataTable dataRecords = new DataTable();
                dataRecords = result.Tables[0];
                dataRecords.Rows[0].Delete();
                dataRecords.AcceptChanges();
                List<PostModel> reqData = new List<PostModel>();
                if (dataRecords.Rows.Count > 0)
                {

                    foreach (DataRow row in dataRecords.Rows)
                    {
                        if (row[0] != null && row[0].ToString().Trim().Length > 0 && list.Any(s => row[0].ToString().Contains(s.NEWSID)))
                        {
                            Console.WriteLine(row[0]);
                            PostModel exceldata = new PostModel();
                            rowexecl++;
                            exceldata.NEWSID = row[0].ToString();
                            exceldata.THUMBNAILIMAGE = row[1].ToString();
                            exceldata.DETAILIMAGE = row[2].ToString();
                            exceldata.TITLE = row[3].ToString();
                            exceldata.URL = row[4].ToString();
                            exceldata.METATITLE = row[5].ToString();
                            exceldata.METAKEYWORD = row[6].ToString();
                            exceldata.METADESCRIPTION = row[7].ToString();
                            exceldata.CREATEDDATE = Convert.ToDateTime(row[8]);
                            exceldata.CREATEDUSER = row[9].ToString();
                            exceldata.CREATEDCUSTOMERID = row[10].ToString();
                            exceldata.TAGS = row[11].ToString();
                            exceldata.LISTCATEGORYID = row[12].ToString();
                            exceldata.LISTCATEGORYNAME = row[13].ToString();


                            String Url = "https://cdn.tgdd.vn/Files/";

                            string Year = exceldata.CREATEDDATE.Year.ToString();
                            string month = exceldata.CREATEDDATE.Month.ToString();
                            if (month.Length == 1)
                            {
                                month = "0" + month;

                            }
                            string day = exceldata.CREATEDDATE.Day.ToString();
                            if (day.Length == 1)
                            {
                                day = "0" + day;

                            }
                            Url += Year + "/" + month + "/" + day + "/" + exceldata.NEWSID + "/";
                            if (!exceldata.DETAILIMAGE.Equals("null") && !exceldata.DETAILIMAGE.Contains("https://"))
                            {

                                Url += exceldata.DETAILIMAGE;
                                string arrListStr = exceldata.DETAILIMAGE.Replace(".jpg", "_300x300.jpg").Replace(".png", "_300x300.png").Replace(".jpeg", "_300x300.jpeg");

                                var namefile = arrListStr;
                                //Tạo thư mục
                                String pathimg = PathProject + "\\NewsFull\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID;
                                System.IO.DirectoryInfo dinew = new DirectoryInfo(pathimg);
                                dinew.Create();
                                // Lấy ảnh từ url vừa tạo ra
                                GetFileFromUrl(pathimg + "\\" + namefile, Url);

                                var a = pathimg + "\\" + namefile;
                                Image image = Image.FromFile(pathimg + "\\" + namefile);
                                //Tạo thư mục
                                String pathimgNews = PathProject + "\\News\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID + "\\";
                                System.IO.DirectoryInfo di = new DirectoryInfo(pathimgNews);
                                di.Create();
                                // chỉnh sửa file về đúng kích thước
                                ResizeAndCompress(image, pathimgNews + namefile, 300, 172);
                                image.Dispose();

                            }
                            if (exceldata.DETAILIMAGE.Equals("null") && !exceldata.THUMBNAILIMAGE.Equals("null") && !exceldata.DETAILIMAGE.Contains("https://"))

                            {

                                // tương tự trên nhưng lấy THUMBNAILIMAGE
                                Url += exceldata.THUMBNAILIMAGE;
                                string arrListStr = exceldata.THUMBNAILIMAGE.Replace(".jpg", "_300x300.jpg").Replace(".png", "_300x300.png").Replace(".jpeg", "_300x300.jpeg");

                                var namefile = arrListStr;

                                String pathimg = PathProject + "\\NewsFull\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID;
                                System.IO.DirectoryInfo dinew = new DirectoryInfo(pathimg);
                                dinew.Create();
                                GetFileFromUrl(pathimg + "\\" + namefile, Url);
                                Image image = Image.FromFile(pathimg + "\\" + namefile);

                                String pathimgNews = PathProject + "\\News\\" + Year + "\\" + month + "\\" + day + "\\" + exceldata.NEWSID + "\\";
                                System.IO.DirectoryInfo di = new DirectoryInfo(pathimgNews);
                                di.Create();

                                ResizeAndCompress(image, pathimgNews + namefile, 300, 172);
                                image.Dispose();


                            }

                            exceldata.Link = Url;
                            reqData.Add(exceldata);
                            DSIdNewsError.Add(exceldata.NEWSID);
                            DSlinkNewsError.Add(exceldata.Link);
                            Console.WriteLine(exceldata.NEWSID);
                        }
                    }

                }

                stream.Close();
                stream.Dispose();
                return 1;

            }
            catch (Exception e)
            {
                stream.Close();
                stream.Dispose();
                Console.WriteLine("File bị lỗi, vui lòng kiểm tra lại file , chú ý dòng " + rowexecl);
                return 2;
            }
        }
    }

}
