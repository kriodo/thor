package exporter

import (
	"github.com/kriodo/thor/office/excel/header"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
	"testing"
)

func TestExporter(t *testing.T) {
	var (
		fileName = "./test/导出测试.xlsx"
	)
	os.Remove(fileName)
	ep, err := NewExporter("测试-1")
	if err != nil {
		t.Log(err)
		return
	}
	// tree结构的表头
	treeHeader := []*header.Header{
		{
			Title:    "子项目名称",
			FieldKey: "subitem_name",
			Children: nil,
		},
		{
			Title:    "调整前",
			FieldKey: "before",
			Children: []*header.Header{
				{
					Title:    "养老",
					FieldKey: "before_old_age",
					Children: []*header.Header{
						{
							Title:    "企业基数",
							FieldKey: "before_old_age_company_base",
							Children: nil,
						},
						{
							Title:    "个人基数",
							FieldKey: "before_old_age_personnel_base",
							Children: nil,
						},
					},
				},
			},
		},
		{
			Title:    "调整后",
			FieldKey: "after",
			Children: []*header.Header{
				{
					Title:    "养老",
					FieldKey: "after_old_age",
					Children: []*header.Header{
						{
							Title:    "企业基数",
							FieldKey: "after_old_age_company_base",
							Children: nil,
						},
						{
							Title:    "个人基数",
							FieldKey: "after_old_age_personnel_base",
							Children: nil,
						},
					},
				},
			},
		},
		{
			Title:    "备注",
			FieldKey: "remark",
			Children: nil,
		},
	}
	// list表头
	var listHeaders []*header.Header
	ep.SetDataStartLine(2)
	err = ep.SetTree(treeHeader).Error()
	if err != nil {
		t.Log(err)
		return
	}
	//var row1s []map[string]interface{}
	//row := make(map[string]interface{})
	//row["subitem_name"] = "测试子项目"
	//row["before_old_age_company_base"] = "99"
	//row1s = append(row1s, row)
	//err = ep.SetMapData(row1s).Error()
	//excelDrop := NewDrop(ep.file, ep.NowSheetName)
	//excelErr := excelDrop.SetDropListV2(&DropListV2Req{
	//	DropList: []string{"" +
	//		"地方简单来说就发了随机发",
	//		"gre人特瑞特尔特额特儿童",
	//		"太热特特1热特务",
	//		"太热特瑞3特让我特务",
	//		"地方简单2来说就发了随机发",
	//		"gre人特瑞特尔特额特儿童",
	//		"太热特特热4特务",
	//		"太热特瑞5特让我特务",
	//		"地方简单来说就54发了随机发",
	//		"gre人特6瑞特尔特额特儿童",
	//		"太热特特7热特务",
	//		"太热特瑞345特让我特务",
	//		"地方简单8来说就发了随机发",
	//		"gre人特瑞3特尔2特额特儿童",
	//		"太热7特特热99特务",
	//		"太热特瑞特让1252我特务",
	//		"地方6简单来说就发了随机发",
	//		"gre人特瑞特尔特5额特儿童",
	//		"太热特特1热4特务",
	//		"太热2特瑞3特让我特务",
	//		"地方简单2来说就发了随机发",
	//		"gre人特瑞2特尔特额特儿童",
	//		"太热特2特热4特务",
	//		"太热特瑞5特让我特务",
	//		"地方简单来2说就54发了随机发",
	//		"gre人特6瑞特尔特额特儿童",
	//		"太热特2特7热特务",
	//		"太热2特瑞345特让我特务",
	//		"地方简单8来说就发了随机发",
	//		"gre人特4瑞特尔2特额特儿童",
	//		"太热特特2热99特务",
	//		"太热3特瑞特让1252我特务",
	//	},
	//	DropKey:     "household_registration_type",
	//	XIndex:      0,
	//	YStartIndex: 6,
	//	YEndIndex:   100,
	//})
	//if excelErr != nil {
	//	t.Log(err)
	//	return
	//}
	//dropKey1, err := ep.SetDropList(ep.File(), &ExcelDropListReq{
	//	XIndex:         1,
	//	YIndex:         3,
	//	YMaxIndex:      10000,
	//	DropSheetName:  "子项目枚举",
	//	DropList:       xdutil.UniqueString(subitemNameList),
	//	IsMax:          true,
	//	QuoteSheetName: sheet,
	//})
	err = ep.SaveAs("./导出.xlsx")
	if err != nil {
		t.Log(err)
		return
	}
	t.Log("------ success")
}

func TestNewExportProcessor2(t *testing.T) {
	os.Remove("./file/b.xlsx")
	ep, err := NewExporter("测试")
	//aFile, err := excelize.OpenFile("a.xlsx")
	//if err != nil {
	//	t.Log(err)
	//	return
	//}
	//defer aFile.Close()
	//ep, err := NewExporterWithFile(aFile)
	if err != nil {
		t.Log(err)
		return
	}

	//header1s := []*header.Header{
	//	{
	//		Title:    "子项目名称",
	//		FieldKey: "suitem_name",
	//		Children: nil,
	//	},
	//	{
	//		Title:    "基础信息",
	//		FieldKey: "before",
	//		Children: []*header.Header{
	//			{
	//				Title:    "姓名",
	//				FieldKey: "user_name",
	//			},
	//			{
	//				Title:    "证件类型",
	//				FieldKey: "user_idcard",
	//			},
	//		},
	//	},
	//	{
	//		Title:    "备注",
	//		FieldKey: "remark",
	//		Children: nil,
	//	},
	//}
	//err = ep.SetTree(header1s).Error()
	//if err != nil {
	//	t.Log(err)
	//	return
	//}
	//var row1s [][]interface{}
	//var row []interface{}
	//row = append(row, "测试子项目")
	//row = append(row, "张三")
	//row = append(row, "100001")
	//row = append(row, "100")
	//row = append(row, "10%")
	//row1s = append(row1s, row)
	//err = ep.SetListData(row1s).Error()
	//if err != nil {
	//	t.Log(err)
	//	return
	//}

	var header1s []*header.Header

	header1s = append(header1s, &header.Header{Id: 1, Pid: 0, Title: "子项目名称", FieldKey: "subitem_name"})
	header1s = append(header1s, &header.Header{Id: 19, Pid: 0, Title: "姓名", Weight: 2, FieldKey: "name"})
	header1s = append(header1s, &header.Header{Id: 20, Pid: 0, Title: "证件号", FieldKey: "id_card"})
	header1s = append(header1s, &header.Header{Id: 2, Pid: 0, Title: "调整前", FieldKey: "before"})
	header1s = append(header1s, &header.Header{Id: 3, Pid: 2, Title: "养老", FieldKey: "before_old_age"})
	header1s = append(header1s, &header.Header{Id: 4, Pid: 3, Title: "企业", FieldKey: "before_old_age_company"})
	header1s = append(header1s, &header.Header{Id: 5, Pid: 4, Title: "企业基数", FieldKey: "before_old_age_company_base"})
	header1s = append(header1s, &header.Header{Id: 6, Pid: 4, Title: "企业比例", Weight: 1, FieldKey: "before_old_age_company_rate"})
	header1s = append(header1s, &header.Header{Id: 7, Pid: 3, Title: "个人", FieldKey: "before_old_age_personnel"})
	header1s = append(header1s, &header.Header{Id: 8, Pid: 7, Title: "个人基数", FieldKey: "before_old_age_personnel_base"})
	header1s = append(header1s, &header.Header{Id: 9, Pid: 7, Title: "个人比例", FieldKey: "before_old_age_personnel_rate"})
	header1s = append(header1s, &header.Header{Id: 10, Pid: 0, Title: "调整后", FieldKey: "after"})
	header1s = append(header1s, &header.Header{Id: 11, Pid: 10, Title: "养老", FieldKey: "after_old_age"})
	header1s = append(header1s, &header.Header{Id: 12, Pid: 11, Title: "企业", FieldKey: "after_old_age_company"})
	header1s = append(header1s, &header.Header{Id: 13, Pid: 12, Title: "企业基数", FieldKey: "after_old_age_company_base"})
	header1s = append(header1s, &header.Header{Id: 14, Pid: 12, Title: "企业比例", FieldKey: "after_old_age_company_rate"})
	header1s = append(header1s, &header.Header{Id: 15, Pid: 11, Title: "个人", FieldKey: "after_old_age_personnel"})
	header1s = append(header1s, &header.Header{Id: 16, Pid: 15, Title: "个人基数", FieldKey: "after_old_age_personnel_base"})
	header1s = append(header1s, &header.Header{Id: 17, Pid: 15, Title: "个人比例", FieldKey: "after_old_age_personnel_rate"})
	header1s = append(header1s, &header.Header{Id: 18, Pid: 0, Title: "备注", Weight: 1, FieldKey: "remark"})
	err = ep.SetList(header1s).Error()
	if err != nil {
		t.Log(err)
		return
	}
	var row1s []map[string]interface{}
	row := make(map[string]interface{})
	row["subitem_name"] = "测试子项目"
	row["name"] = "测试姓名"
	row["before_old_age"] = "xxxxxx"
	row["after_old_age_personnel_rate"] = "20%"
	row["remark"] = "这是一个备注"
	row["after_old_age_personnel_rate"] = FormatAmountFloatIncludeMinus1("-1111")
	row1s = append(row1s, row)
	err = ep.SetMapData(row1s).Error()
	if err != nil {
		t.Log(err)
		return
	}

	//// 设置样式
	//styleStr := "###0;[Red]-###0"
	//styleId, err := ep.file.NewStyle(&excelize.Style{CustomNumFmt: &styleStr})
	//var aa = make(map[string]int)
	//aa["after_old_age_personnel_rate"] = styleId
	//ep.SetRowStyle(aa)

	ep.AddSheet("测试2")
	ep.SetStartIndex(3)
	var header2s []*header.Header
	header2s = append(header2s, &header.Header{Pkey: "", Title: "子项目名称", FieldKey: "subitem_name2"})
	header2s = append(header2s, &header.Header{Pkey: "", Title: "姓名", Weight: 2, FieldKey: "name2"})
	header2s = append(header2s, &header.Header{Pkey: "", Title: "证件号", FieldKey: "id_card2"})
	header2s = append(header2s, &header.Header{Pkey: "", Title: "调整前", FieldKey: "before2"})
	header2s = append(header2s, &header.Header{Pkey: "before2", Title: "养老", FieldKey: "before_old_age2"})
	header2s = append(header2s, &header.Header{Pkey: "before_old_age2", Title: "企业", FieldKey: "before_old_age_company2"})
	header2s = append(header2s, &header.Header{Pkey: "before_old_age2", Title: "个人", FieldKey: "before_old_age_personal2"})
	header2s = append(header2s, &header.Header{Pkey: "", Title: "调整后", FieldKey: "after2"})
	header2s = append(header2s, &header.Header{Pkey: "after2", Title: "养老", FieldKey: "after_old_age2"})
	header2s = append(header2s, &header.Header{Pkey: "after_old_age2", Title: "企业", FieldKey: "after_old_age_company2"})
	header2s = append(header2s, &header.Header{Pkey: "after_old_age2", Title: "个人", FieldKey: "after_old_age_personal2"})
	err = ep.SetListV2(header2s).Error()

	var row2s []map[string]interface{}
	row2 := make(map[string]interface{})
	row2["subitem_name2"] = "测试子项目2"
	row2["name2"] = "测试姓名2"
	row2["id_card2"] = "123123123213"
	row2["before_old_age_company2"] = "20%"
	row2s = append(row2s, row2)
	err = ep.SetMapData(row2s).Error()
	if err != nil {
		t.Log(err)
		return
	}
	err = ep.SaveAs("b.xlsx")
	if err != nil {
		t.Log(err)
		return
	}
	t.Log("------ success")
}

func TestNewExportProcessor3(t *testing.T) {
	os.Remove("b.xlsx")
	aFile, err := excelize.OpenFile("./file/a.xlsx")
	if err != nil {
		t.Log(err)
		return
	}
	defer aFile.Close()
	ep, err := NewExporterWithFile(aFile)
	if err != nil {
		t.Log(err)
		return
	}
	ep.NowSheetName = "Sheet1"

	ep.SetRowStartLine(7)

	var row1s [][]interface{}
	var row []interface{}
	row = append(row, "测试子项目")
	row = append(row, "张三")
	row = append(row, "100001")
	row = append(row, "100")
	row = append(row, "10%")
	row1s = append(row1s, row)
	err = ep.SetListData(row1s).Error()
	if err != nil {
		t.Log(err)
		return
	}

	err = ep.SaveAs("b.xlsx")
	if err != nil {
		t.Log(err)
		return
	}
	t.Log("------ success")
}

// FormatAmountFloatIncludeMinus1 格式化金额 100 => 100.00
// 如果是-1, 返回原值
func FormatAmountFloatIncludeMinus1(amount string) interface{} {
	if amount == "" {
		return ""
	}

	numberFloat, _ := strconv.ParseFloat(amount, 64)
	return numberFloat
}
