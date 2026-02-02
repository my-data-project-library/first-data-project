let
    源 = Excel.CurrentWorkbook(){[Name="表1"]}[Content],
    日期初步清理 = Table.AddColumn(源,"日期半清理",each if Type.Is(Value.Type(_[订单日期]),type datetime) or [订单日期]=null then [订单日期] else DateTime.FromText(Text.Replace(Text.From(_[订单日期]),Text.End(Text.From(_[订单日期]),2),"00"))),
    日期清理完成 = Table.AddColumn(日期初步清理,"日期全清理",each if  [日期半清理]<>null then [日期半清理] else Table.Group(日期初步清理,"用户ID",{"去空日期",each List.Sort(List.RemoveNulls([日期半清理]))}){[用户ID=[用户ID]]}[去空日期]{0}),
    转为标准日期 = Table.AddColumn(日期清理完成,"标准日期",each DateTime.ToText([日期全清理],"yyyy-MM-dd")),
    过滤零元记录 = Table.SelectRows(转为标准日期,each [支付金额]<>0),
    引用商品表格 = Table.AddColumn(过滤零元记录,"引用表格", each
let
    源 = Excel.CurrentWorkbook(){[Name="表5"]}[Content],
    过滤零元记录 = Table.SelectRows(源,each [单品单价]<>0),
    表分组 = Table.Group(过滤零元记录,"订单ID",{"总金额",each List.Sum(Table.AddColumn(_,"总金额",each [购买数量]*[单品单价])[总金额])})
in  表分组),
    检验是否异常 = Table.AddColumn(引用商品表格,"是否正常",each if [支付金额]=[引用表格]{[订单ID=[订单ID]]}[总金额] then "金额正常" else "金额异常"),
    处理发货仓库 = Table.AddColumn(检验是否异常,"仓库去除空",each try if [发货仓库]<>null then [发货仓库] else // User_Info
let
    源 = Excel.CurrentWorkbook(){[Name="表3"]}[Content]{[用户ID=[用户ID]]}[所在城市]&"仓"
in
    源 otherwise "默认仓"),
    删除多余的列 = Table.RemoveColumns(处理发货仓库,{"订单日期", "发货仓库"}),
    完善数据类型 = Table.TransformColumns(删除多余的列,{{"标准日期",each Date.From(_)},{"支付金额",each Number.From(_)}},each Text.From(_))
in
    完善数据类型




let
    源 = Table.NestedJoin(合并1, {"用户ID"}, User_Info, {"用户ID"}, "User_Info", JoinKind.LeftOuter),
    提取年月 = Table.AddColumn(源,"年月",each Date.ToText([标准日期],"yyyy-M")),
    客户等级 = Table.AddColumn(提取年月,"等级",each if [支付金额]>=500 then "高客单" else if [支付金额]>=100 then "中客单" else "低客单"),
    复购情况 = Table.AddColumn(Table.AddColumn(客户等级,"是否复购",each if Table.RowCount([Goods_Detail])>=2 then "是" else "否"),"是否为优质订单",each if [支付金额]>=200 and [订单状态]="已完成" and [是否复购]="是" then "优质" else "普通"),
    继续分组 = Table.Group(复购情况,{"仓库去除空","等级","年月"},{{"分组",each _},{"订单数",each List.Count(List.Distinct([订单ID]))},{"支付总金额",each List.Sum([支付金额])},{"平均支付金额",each Number.Round(List.Average([支付金额]),2)},{"复购订单数",each List.Count(List.Select([是否复购],each _="是"))},{"优质订单占比",each List.Count(List.Select([是否为优质订单],each _="优质"))/List.Count(List.Distinct([订单ID]))}}),
    更改类型 = Table.TransformColumnTypes(继续分组,{{"优质订单占比",Percentage.Type}}),
    表格排序 = Table.Sort(更改类型,{{"支付总金额",1},{"订单数",1}}),
    展示部分 = Table.SelectColumns(表格排序,{"分组","订单数","支付总金额","平均支付金额","复购订单数","优质订单占比"})
in
    展示部分



