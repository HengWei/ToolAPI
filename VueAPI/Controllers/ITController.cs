using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using VueAPI.Models;
using System.Web.Http.Cors;
using System.Configuration;
using System.Net.Http.Headers;

namespace VueAPI.Controllers
{
    [EnableCors(origins: "http://192.168.1.145, http://localhost:1337", headers: "*", methods: "*")]
    public class ITController : ApiController
    {
        [HttpGet]
        public string Test(string id)
        {
            return id;
        }


        [HttpGet]
        public HttpResponseMessage DownloadFile(string f)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/") + string.Format(@"{0}{1}", ConfigurationManager.AppSettings["DownloadPath"], f);
            string fileName = Path.GetFileName(filePath);

            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    byte[] bytes = new byte[file.Length];
                    file.Read(bytes, 0, (int)file.Length);
                    ms.Write(bytes, 0, (int)file.Length);

                    HttpResponseMessage response = new HttpResponseMessage();
                    response.Content = new ByteArrayContent(bytes.ToArray());
                    response.Content.Headers.Add("x-filename", fileName);
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                    response.Content.Headers.ContentDisposition.FileName = fileName;
                    response.StatusCode = HttpStatusCode.OK;
                    return response;
                }
            }
        }


        [HttpPost]
        public string UploadPackingSerial()
        {
            try
            {
                var uploadResult = ImportExcel(HttpContext.Current.Request);

                if (!uploadResult.IsSuccess)
                {
                    throw new Exception(uploadResult.Message);
                }

                var inputFileName = HttpContext.Current.Request.Files[0].FileName;

                var fileName = inputFileName.Insert(inputFileName.LastIndexOf("."), "_NEW");                 

                string outputFile = string.Format(HttpContext.Current.Server.MapPath("~/") + string.Format(@"App_Data\Download\{0}", fileName));


                using (XLWorkbook workbook = new XLWorkbook(uploadResult.File))
                {
                    IXLWorksheet worksheet = workbook.Worksheet(1);
                    // 定義資料起始、結束 Cell
                    var firstCell = worksheet.FirstCellUsed();
                    var lastCell = worksheet.LastCellUsed();

                    // 使用資料起始、結束 Cell，來定義出一個資料範圍
                    var data = worksheet.Range(firstCell.Address, lastCell.Address);

                    // 將資料範圍轉型
                    var table = data.AsTable();
                    int i = 0;


                    string preBoxNo = string.Empty;
                    //目前箱號
                    int boxNo = 0;

                    foreach (var row in table.Rows())
                    {
                        i++;

                        int boxValue = -1;
                        //標題列跳過
                        if (i == 1)
                        {
                            continue;
                        }

                        var strPreBoxNo = row.Cell("B").Value.ToString();

                        if (string.IsNullOrEmpty(strPreBoxNo))
                        {
                            row.Cell("B").SetValue(preBoxNo);
                        }
                        else
                        {
                            preBoxNo = strPreBoxNo.Trim().ToUpper();
                            boxNo = 0;
                        }

                        var strBoxValue = row.Cell("G").Value.ToString();

                        if (!string.IsNullOrEmpty(strBoxValue))
                        {
                            if (int.TryParse(strBoxValue, out int newBoxValue))
                            {
                                boxValue = newBoxValue;

                                if (boxValue > 0)
                                {
                                    boxNo++;
                                }
                            }
                        }

                        var strBoxNo = row.Cell("D").Value.ToString();

                        if (string.IsNullOrEmpty(strBoxNo))
                        {

                        }
                        else
                        {
                            if (int.TryParse(strBoxNo, out int newBoxNo))
                            {
                                boxNo = newBoxNo;

                            }
                        }

                        row.Cell("D").SetValue(boxNo.ToString());


                        if (boxValue == 0)
                        {


                        }
                        else if (boxValue > 0)
                        {
                            boxNo = (boxNo + boxValue - 1);
                        }
                        else
                        {
                            Console.WriteLine("第 {0} 列，BOX_VALUUE 資料異常，請確認", i);
                        }

                        row.Cell("E").SetValue((boxNo).ToString());

                        //BOX Value
                        if (string.IsNullOrEmpty(row.Cell("G").Value.ToString()))
                        {
                            row.Cell("G").SetValue(0);
                        }

                        //CUFT
                        if (string.IsNullOrEmpty(row.Cell("I").Value.ToString()))
                        {
                            row.Cell("I").SetValue(0);
                        }

                        //NW
                        if (string.IsNullOrEmpty(row.Cell("J").Value.ToString()))
                        {
                            row.Cell("J").SetValue(0);
                        }

                        //GW
                        if (string.IsNullOrEmpty(row.Cell("K").Value.ToString()))
                        {
                            row.Cell("K").SetValue(0);
                        }
                    }

                    workbook.SaveAs(outputFile);
                }

                return fileName;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }            
        }


        public ImportFileViewModel ImportExcel(HttpRequest _Request)
        {
            ImportFileViewModel result = new ImportFileViewModel();

            if (_Request.Files.Count > 0)
            {
                // 取得檔案
                HttpPostedFile file = _Request.Files[0];
                // 檢查附檔名
                string[] ext = { ".xlsx" };
                if (!ext.Any(o => Path.GetExtension(file.FileName).ToLower() == o))
                {
                    result.Message = "檔案類型錯誤，只允許副檔名為xlsx。";
                    result.IsSuccess = false;
                    return result;
                }

                //web.config 路徑                
                string configPath = HttpContext.Current.Server.MapPath("~/") + @"App_Data\Upload";

                string graphType = _Request.Form["type"];
                string fileName = string.Format("{0}", DateTime.Now.ToString("yyyyMMddHHmmss"));

                // 設定資料夾
                string absoluteDirPath = $"{configPath}/{graphType}";
                if (!Directory.Exists(absoluteDirPath))
                    Directory.CreateDirectory(absoluteDirPath);

                // 設定儲存路徑
                string absoluteFilePath = absoluteDirPath + $"/{fileName}{Path.GetExtension(file.FileName)}";


                // 儲存
                file.SaveAs(absoluteFilePath);
                result.File = absoluteFilePath;
                return result;
            }
            else
            {
                result.Message = "尚未選擇上傳檔案。";
                result.IsSuccess = false;
                return result;
            }
        }

        [HttpGet]
        public IEnumerable<GuidViewModel> GetProductGroup()
        {
            return new Service.Service().GetGroupData();
        }

        [HttpGet]
        public IEnumerable<GuidViewModel> GetProductCategory(Guid GroupId)
        {
            return new Service.Service().GetCategoryData(GroupId);
        }

        [HttpPost]
        public ResultModel InsertProductName(ProductNameModel _Insert)
        {
            if(_Insert.GroupId == Guid.Empty)
            {
                return new ResultModel(){ 
                    IsSucess = false,
                    Message = "空白的Group Id"
                };
            }

            if (_Insert.CategoryId == Guid.Empty)
            {
                return new ResultModel()
                {
                    IsSucess = false,
                    Message = "空白的Category Id"
                };
            }

            if (string.IsNullOrEmpty(_Insert.ProductName))
            {
                return new ResultModel()
                {
                    IsSucess = false,
                    Message = "空白的Product Name"
                };
            }

            return new Service.Service().InsertSuggestName(_Insert.GroupId, _Insert.CategoryId, _Insert.ProductName);
        }

        [HttpPost]
        public ResultModel FindModel()
        {
            try
            {
                var uploadResult = ImportExcel(HttpContext.Current.Request);

                if (!uploadResult.IsSuccess)
                {
                    throw new Exception(uploadResult.Message);
                }

                var inputFileName = HttpContext.Current.Request.Files[0].FileName;

                var fileName = inputFileName.Insert(inputFileName.LastIndexOf("."), "_NEW");

                string outputFile = string.Format(HttpContext.Current.Server.MapPath("~/") + string.Format(@"App_Data\Download\{0}", fileName));

                Service.Service service = new Service.Service();

                using (XLWorkbook workbook = new XLWorkbook(uploadResult.File))
                {
                    IXLWorksheet worksheet = workbook.Worksheet(1);
                    // 定義資料起始、結束 Cell
                    var firstCell = worksheet.FirstCellUsed();
                    var lastCell = worksheet.LastCellUsed();

                    // 使用資料起始、結束 Cell，來定義出一個資料範圍
                    var data = worksheet.Range(firstCell.Address, lastCell.Address);

                    // 將資料範圍轉型
                    var table = data.AsTable();
                    int partNoColum = 0;
                    var firstRow = table.Rows().FirstOrDefault();

                    //找PartNo位置
                    var existColumnNumber = firstRow.RangeAddress.NumberOfCells;

                    for (int k = 1; k <= existColumnNumber; k++)
                    {
                        var strColumn = firstRow.Cell(k).Value.ToString().Trim();

                        if (!string.IsNullOrEmpty(strColumn) && strColumn.Equals("Part No", StringComparison.OrdinalIgnoreCase))
                        {
                            partNoColum = k;
                        }
                    }

                    if (partNoColum == 0)
                    {
                        return new ResultModel() {
                            IsSucess = false,
                            Message = "請於第一列將[料號]欄位標記改為[Part No]"
                        };
                    }

                    //建立新位置
                    for (int j = 0; j < 3; j++)
                    {
                        firstRow.Cell(existColumnNumber + 1 + j * 3).SetValue(string.Format("Car Brand({0})", j + 1));
                        firstRow.Cell(existColumnNumber + 2 + j * 3).SetValue(string.Format("Car Model({0})", j + 1));
                        firstRow.Cell(existColumnNumber + 3 + j * 3).SetValue(string.Format("Car Year({0})", j + 1));
                    }


                    //資料處理
                    foreach (var row in table.Rows())
                    {
                        var ordePTNO = row.Cell(partNoColum).Value.ToString().Trim().ToUpper();

                        var modelData = service.GetCarModel(ordePTNO);

                        int c = 0;

                        foreach (var item in modelData.Take(3))
                        {
                            //4
                            row.Cell(existColumnNumber + 1 + c * 3).SetValue(item.Make);
                            //5
                            row.Cell(existColumnNumber + 2 + c * 3).SetValue(item.Model);
                            //6
                            row.Cell(existColumnNumber + 3 + c * 3).SetValue(string.Format("{0}-{1}", item.BeginYear, item.EndYear));

                            c++;
                        }
                    }

                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(outputFile);
                }

                return new ResultModel()
                {
                    IsSucess = true,
                    Message = fileName
                };

            }
            catch (Exception ex)
            {
                return new ResultModel()
                {
                    IsSucess = false,
                    Message = ex.Message
                };
            }


            
        }

        [HttpGet]
        public string ExportPDList(DateTime Start, DateTime End)
        {
            var fileName = DateTime.Now.ToString(@"yyyyMMddhhmmss");

            string outputFile = string.Format(HttpContext.Current.Server.MapPath("~/") + string.Format(@"App_Data\Download\{0}.xlsx", fileName));

            Service.Service service = new Service.Service();

            var data = service.GetPDList(Start, End.AddDays(1).AddSeconds(-1));

            var groupList = data.GroupBy(x => x.GroupCode).Select(x => x.Key).OrderBy(x=>x);

            using (XLWorkbook workbook = new XLWorkbook())
            {
                foreach (var item in groupList)
                {
                    IXLWorksheet worksheet = workbook.AddWorksheet(item);

                    var itemGroup = data.Where(x => string.Equals(x.GroupCode, item, StringComparison.OrdinalIgnoreCase));

                    int i = 0;

                    int j = 0;

                    worksheet.Row(++i).Cell(++j).SetValue("代表料號");
                    worksheet.Row(i).Cell(++j).SetValue("PDName");
                    worksheet.Row(i).Cell(++j).SetValue("Make");
                    worksheet.Row(i).Cell(++j).SetValue("Model");
                    worksheet.Row(i).Cell(++j).SetValue("BeginYear");
                    worksheet.Row(i).Cell(++j).SetValue("EndYear");
                    worksheet.Row(i).Cell(++j).SetValue("報價單號");


                    foreach (var gItem in itemGroup.OrderBy(x=>x.PDName))
                    {
                        j = 0;

                        worksheet.Row(++i).Cell(++j).SetValue(gItem.MajorRefNo);
                        worksheet.Row(i).Cell(++j).SetValue(gItem.PDName);
                        worksheet.Row(i).Cell(++j).SetValue(gItem.Make??string.Empty);
                        worksheet.Row(i).Cell(++j).SetValue(gItem.Model??string.Empty);
                        worksheet.Row(i).Cell(++j).SetValue(gItem.BeginYear??string.Empty);
                        worksheet.Row(i).Cell(++j).SetValue(gItem.EndYear??string.Empty);
                        worksheet.Row(i).Cell(++j).SetValue(gItem.QueryNo??string.Empty);
                    }

                    worksheet.Columns().AdjustToContents();
                }

                workbook.SaveAs(outputFile);

                return outputFile;
            }
        }

        [HttpGet]
        public string ExportRecommendList(string BuyerCode)
        {
            Service.Service _Service = new Service.Service();

            var source = _Service.GetOrderInfo(1);

            List<ExportRelation> exportList = new List<ExportRelation>();

            //掃描來源清單
            foreach (var item in source)
            {
                //找出在關聯表裡的資料
                var target = _Service.GetRelationList(item.OrderPTNO);


                ExportRelation temp = new ExportRelation();

                temp.SourcePTNo = item.OrderPTNO;
                temp.SourcePDName = item.PDName;
                temp.SourceCategoryName = item.CategoryNameEN;
                temp.SourceOrderNo = item.OrderNo;

                if (target.Count()==0)
                {
                    continue;
                }
                //有資料的話
                else
                {
                    foreach (var tItem in target)
                    {
                        if (string.Equals(tItem.RefNo1, item.OrderPTNO, StringComparison.OrdinalIgnoreCase))
                        {
                            temp.TargetPTNo = tItem.RefNo2;
                        } 
                        else 
                        {
                            temp.TargetPTNo = tItem.RefNo1;
                        }

                        if(source.Any(x=>x.OrderPTNO==temp.TargetPTNo))
                        {
                            continue;
                        }


                        var tData = _Service.GetRelationInfo(temp.TargetPTNo);

                        temp.TargetCategoryName = tData.CategoryNameEN;
                        temp.TargetPDName = tData.PDName;
                        temp.Times = tItem.Times;
                        temp.Rate = tItem.Rate;

                        exportList.Add(temp);

                        break;
                    }
                    

                    
                }


            }

            var fileName = DateTime.Now.ToString(@"yyyyMMddhhmmss");

            string outputFile = string.Format(HttpContext.Current.Server.MapPath("~/") + string.Format(@"App_Data\Download\{0}.xlsx", fileName));

            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.AddWorksheet("推薦清單");

                int i = 0;

                int j = 0;

                worksheet.Row(++i).Cell(++j).SetValue("推薦料號");
                worksheet.Row(i).Cell(++j).SetValue("推薦品名");
                worksheet.Row(i).Cell(++j).SetValue("推薦類別");
                worksheet.Row(i).Cell(++j).SetValue("根據料號");
                worksheet.Row(i).Cell(++j).SetValue("根據品名");
                worksheet.Row(i).Cell(++j).SetValue("根據料號類別");
                worksheet.Row(i).Cell(++j).SetValue("根據訂單號");
                worksheet.Row(i).Cell(++j).SetValue("購買組合次數");
                worksheet.Row(i).Cell(++j).SetValue("購買組合比例");


                foreach (var item in exportList)
                {
                    j = 0;
                    worksheet.Row(++i).Cell(++j).SetValue(item.TargetPTNo);
                    worksheet.Row(i).Cell(++j).SetValue(item.TargetPDName);
                    worksheet.Row(i).Cell(++j).SetValue(item.TargetCategoryName);
                    worksheet.Row(i).Cell(++j).SetValue(item.SourcePTNo);
                    worksheet.Row(i).Cell(++j).SetValue(item.SourcePDName);
                    worksheet.Row(i).Cell(++j).SetValue(item.SourceCategoryName);
                    worksheet.Row(i).Cell(++j).SetValue(item.SourceOrderNo);
                    worksheet.Row(i).Cell(++j).SetValue(item.Times);
                    worksheet.Row(i).Cell(++j).SetValue(item.Rate);


                }
                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(outputFile);

                return outputFile;
            }

        }
    }
}
