using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using VueAPI.Models;
using Dapper;

namespace VueAPI.Service
{

    public class Service
    {
        protected static readonly string InternalContext = ConfigurationManager.ConnectionStrings["NKMT"].ConnectionString;

        public IEnumerable<GuidViewModel> GetGroupData()
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.Query<GuidViewModel>(@"
                    SELECT GroupId AS Value, GroupCode+'.'+GroupNameEN AS Text
                    FROM dbo.CO_Group WITH(NOLOCK) 
                    WHERE IsDisabled = 0
                    ORDER BY GroupCode;"
                    );
            }
        }

        public IEnumerable<GuidViewModel> GetCategoryData(Guid _GroupId)
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.Query<GuidViewModel>(@"
                    SELECT CategoryId AS Value, CategoryCode+'.'+ CategoryNameEN AS Text
                    FROM dbo.CO_Category WITH(NOLOCK) 
                    WHERE IsDisabled = 0
                    AND GroupId = @GroupId
                    ORDER BY CategoryCode;",
                    new
                    {
                        GroupId = _GroupId
                    });
            }
        }

        public ResultModel InsertSuggestName(Guid _GroupId, Guid _CategoryId, string _SuggetstedName)
        {
            try
            {
                using (var con = new SqlConnection(InternalContext))
                {
                    var data = con.QuerySingleOrDefault<int>(@"
                    SELECT COUNT(*)
                    FROM dbo.CO_SuggestedProductName
                    WHERE SuggestedName = @SuggestedName"
                    , new
                    {

                        SuggestedName = _SuggetstedName
                    });

                    if (data > 1)
                    {
                        return new ResultModel()
                        {
                            IsSucess = false,
                            Message = "已有重複的SuggestedName"
                        };
                    }

                    con.Execute(@"
                    INSERT CO_SuggestedProductName ([GroupId], [CategoryId], [SuggestedName], [CreationTime])
                    VALUES(@GroupId, @CategoryId, @SuggestedName, GETDATE());",
                    new
                    {
                        GroupId = _GroupId,
                        CategoryId = _CategoryId,
                        SuggestedName = _SuggetstedName
                    });
                }

                return new ResultModel()
                {
                    IsSucess = true,
                    Message = string.Empty
                };
            }
            catch (Exception ex)
            {
                return new ResultModel()
                {
                    IsSucess = true,
                    Message = ex.Message
                };
            }
        }

        public IEnumerable<CarModel> GetCarModel(string OrderPTNO)
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.Query<CarModel>(@"
                       SELECT ck.Make, cm.CarModelName AS Model, mf.BeginYear, mf.EndYear
                       FROM dbo.PD_ModelFit AS mf WITH(NOLOCK)
                       LEFT JOIN dbo.CO_CarModel AS cm WITH(NOLOCK) ON mf.CarModelId = cm.CarModelId
                       LEFT JOIN dbo.CO_CarMake AS ck WITH(NOLOCK) ON cm.MakeABB = ck.MakeABB
					   LEFT JOIN dbo.PD_MakeFit AS mk ON mk.PDMainId = mf.PDMainId AND cm.MakeABB = mk.MakeABB
                       WHERE mf.PDMainId IN (SELECT PDMainId 
                       FROM dbo.PD_RefNo WITH(NOLOCK)
                       WHERE RefNo = @RefNo)
					   AND mk.PDMainId IS NOT NULL;",
                       new
                       {
                           RefNo = OrderPTNO
                       });
            }
        }

        public IEnumerable<PDList> GetPDList(DateTime _Start, DateTime _End)
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.Query<PDList>(@"
                       SELECT pm.MajorRefNo, pm.PDName, g.GroupCode, p.QueryNo, m.Make, m.Model, m.BeginYear, m.EndYear
                       FROM PD_Main AS pm
                       LEFT JOIN dbo.CO_Category AS c ON pm.CategoryId = c.CategoryId
                       LEFT JOIN dbo.CO_Group AS g ON g.GroupId = c.GroupId
                       LEFT JOIN (
                       SELECT r.PDMainId ,qm.QueryNo, ROW_NUMBER() OVER(PARTITION by r.pdMainId ORDER BY qm.CreationTime ASC) AS '#row'
                       FROM TSR_QuotesDetail AS qd
                       LEFT JOIN dbo.PD_RefNo AS r ON qd.OrderPTNO = r.RefNo
                       LEFT JOIN dbo.TSR_QuotesMain AS qm ON qd.QuotesMainId = qm.QuotesMainId
                       ) AS p ON p.PDMainId = pm.PDMainId AND p.#row = 1
                       LEFT JOIN (
                       SELECT mf.PDMainId ,ck.Make, cm.CarModelName AS Model, mf.BeginYear, mf.EndYear, ROW_NUMBER() OVER(PARTITION BY mf.PDMainId ORDER BY mk.Sort) AS '#row'
                       FROM dbo.PD_ModelFit AS mf WITH(NOLOCK)
                       LEFT JOIN dbo.CO_CarModel AS cm WITH(NOLOCK) ON mf.CarModelId = cm.CarModelId
                       LEFT JOIN dbo.CO_CarMake AS ck WITH(NOLOCK) ON cm.MakeABB = ck.MakeABB
					   LEFT JOIN dbo.PD_MakeFit AS mk ON mk.PDMainId = mf.PDMainId AND cm.MakeABB = mk.MakeABB
                       ) AS m ON m.PDMainId = pm.PDMainId AND m.#row = 1                        
                       WHERE pm.CreationTime BETWEEN @start AND @end
                       AND pm.IsDeleted = 0;",
                       new
                       {
                           start = _Start,
                           end =_End
                       });
            }
        }

        public IEnumerable<OrderRefNo> GetOrderInfo(string _BuyerCode)
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.Query<OrderRefNo>(@"
                       SELECT *
                       FROM (
                       SELECT qd.OrderPTNO, qm.OrderNo, pm.PDName, c.CategoryCode, c.CategoryNameEN, ROW_NUMBER() OVER(PARTITION BY qd.OrderPTNo ORDER BY qm.OrderingDate DESC) AS '#row'
                       FROM TSR_QuotesDetail AS qd
                       LEFT JOIN TSR_QuotesMain AS qm ON qd.QuotesMainId = qm.QuotesMainId
                       LEFT JOIN BYR_Main AS bm ON bm.BYRMainId = qm.BYRMainId
                       LEFT JOIN dbo.PD_RefNo AS r ON r.RefNo = qd.OrderPTNO
                       LEFT JOIN dbo.PD_Main AS pm ON pm.PDMainId = r.PDMainId
                       LEFT JOIN dbo.CO_Category AS c ON pm.CategoryId = c.CategoryId
                       WHERE qm.OrderNo IS NOT NULL
                       AND qm.OrderingDate > '2019-01-01 00:00:000'
                       AND bm.BuyerCode = @BuyerCode) AS data
                       WHERE data.#row = 1
                       ",
                       new
                       {
                           BuyerCode = _BuyerCode
                       });
            }
        }


        public IEnumerable<OrderRefNo> GetOrderInfo(int _Test)
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.Query<OrderRefNo>(@"
                       SELECT *
                       FROM (
                       SELECT qd.OrderPTNO, ISNULL(qm.OrderNo, qm.QueryNo) AS OrderNo, pm.PDName, c.CategoryCode, c.CategoryNameEN, ROW_NUMBER() OVER(PARTITION BY qd.OrderPTNo ORDER BY qm.OrderingDate DESC) AS '#row'
                       FROM TSR_QuotesDetail AS qd
                       LEFT JOIN TSR_QuotesMain AS qm ON qd.QuotesMainId = qm.QuotesMainId
                       LEFT JOIN BYR_Main AS bm ON bm.BYRMainId = qm.BYRMainId
                       LEFT JOIN dbo.PD_RefNo AS r ON r.RefNo = qd.OrderPTNO
                       LEFT JOIN dbo.PD_Main AS pm ON pm.PDMainId = r.PDMainId
                       LEFT JOIN dbo.CO_Category AS c ON pm.CategoryId = c.CategoryId
                       WHERE QueryNo = 'IXGM220107PI1') AS data
                       WHERE data.#row = 1");
            }
        }

        public OrderRefNo GetRelationInfo(string _OrderPTNo)
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.QuerySingleOrDefault<OrderRefNo>(@"
                       SELECT r.RefNo AS OrderPTNO, pm.PDName, c.CategoryCode, c.CategoryNameEN
                       FROM dbo.PD_RefNo AS r
                       LEFT JOIN dbo.PD_Main AS pm ON pm.PDMainId = r.PDMainId
                       LEFT JOIN dbo.CO_Category AS c ON pm.CategoryId = c.CategoryId
                       WHERE RefNo = @RefNo", new {
                       RefNo = _OrderPTNo
                });
            }
        }

        public IEnumerable<Relation> GetRelationList(string _OrderPTNO)
        {
            using (var con = new SqlConnection(InternalContext))
            {
                return con.Query<Relation>(@"
                       SELECT * 
                       FROM AnalysisRelation
                       WHERE (RefNo1 = @OrderPTNO OR RefNo2 = @OrderPTNO)
                       AND TotalTime > 1
                       ORDER BY TotalTime DESC, Rate DESC;",
                       new
                       {
                           OrderPTNO = _OrderPTNO
                       });
            }
        }
    }


}