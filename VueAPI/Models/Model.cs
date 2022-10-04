using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace VueAPI.Models
{
    public class Model
    {
    }

    public class ImportFileViewModel
    {
        public bool IsSuccess { get; set; } = true;

        public string Message { get; set; } = string.Empty;

        public string File { get; set; } = string.Empty;
    }

    public class GuidViewModel
    {
        public Guid Value { get; set; }

        public string Text { get; set; }
    }

    public class ProductNameModel
    {
        public Guid GroupId { get; set; }

        public Guid CategoryId { get; set; }

        public string ProductName { get; set; }
    }

    public class ResultModel
    {
        public bool IsSucess { get; set; }

        public string Message { get; set; }
    }

    public class CarModel
    {
        public string Make { get; set; }

        public string Model { get; set; }

        public int? BeginYear { get; set; }

        public int? EndYear { get; set; }
    }

    public class PDList
    {
        public string MajorRefNo { get; set; }

        public string PDName { get; set; }

        public string GroupCode { get; set; }

        public string QueryNo { get; set; }

        public string Make { get; set; }

        public string Model { get; set; }

        public string BeginYear { get; set; }

        public string EndYear { get; set; }
    }

    public class OrderRefNo
    {
        public string OrderNo { get; set; }

        public string OrderPTNO { get; set; }

        public string PDName { get; set; }

        public string CategoryCode { get; set; }

        public string CategoryNameEN { get; set; }
    }

    public class Relation
    {
        public string RefNo1 { get; set; }

        public string RefNo2 { get; set; }

        public int Times { get; set; }

        public int TotalTime { get; set; }

        public double Rate { get; set; }
    }

    public class ExportRelation
    {
        public string SourcePTNo { get; set; }

        public string SourcePDName { get; set; }

        public string SourceCategoryName { get; set; }

        public string SourceOrderNo { get; set; }

        public string TargetPTNo { get; set; }

        public string TargetPDName { get; set; }

        public string TargetCategoryName { get; set; }

        public int Times { get; set; }

        public double Rate { get; set; }
    }

    public class ToolReport
    {
        public List<ToolReportItem> data { get; set; }
    }

    public class ToolReportItem
    {
        public string _id { get; set; }

        public int count { get; set; }
    }
}