package importer

import (
	"github.com/kriodo/thor/office/excel/header"
	"github.com/xuri/excelize/v2"
	"testing"
)

func TestNewImporter(t *testing.T) {
	fileName := "../test/导出测试.xlsx"
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		t.Log(err)
		return
	}
	ir := NewImporter(f)
	err = ir.Error()
	if err != nil {
		t.Log(err)
		return
	}
	headers1 := []*header.Header{
		{Title: "姓名", FieldKey: "user_name"},
		{Title: "户籍", FieldKey: "household_registration"},
		{Title: "户口所在城市", FieldKey: "household_registration_city"},
		{Title: "工作城市", FieldKey: "work_city"},
		{Title: "子项目名称", FieldKey: "subitem_name"},
		{Title: "调整前", FieldKey: "before", Children: []*header.Header{
			{Title: "养老", FieldKey: "before_old_age", Children: []*header.Header{
				{Title: "企业基数", FieldKey: "before_old_age_company_base"},
				{Title: "企业比例", FieldKey: "before_old_age_company_rate"},
				{Title: "个人基数", FieldKey: "before_old_age_personnel_base"},
				{Title: "个人比例", FieldKey: "before_old_age_personnel_rate"},
			}}}},
		{Title: "调整后", FieldKey: "after", Children: []*header.Header{
			{Title: "养老", FieldKey: "after_old_age", Children: []*header.Header{
				{Title: "企业基数", FieldKey: "after_old_age_company_base"},
				{Title: "企业比例", FieldKey: "after_old_age_company_rate"},
				{Title: "个人基数", FieldKey: "after_old_age_personnel_base"},
				{Title: "个人比例", FieldKey: "after_old_age_personnel_rate"},
			}}}},
		{Title: "备注", FieldKey: "remark"},
		//{Title: "测试", FieldKey: "demo", Import: header.ImportAttr{IsRequired: true}},
	}
	err = ir.SetHeaderTree(headers1).Error()
	if err != nil {
		t.Log(err)
		return
	}
	dataList := ir.GetRows()
	for i, v := range dataList {
		t.Logf("------- 序号：%d ------------  %s", i, ir.PrintRows(v))
	}

	err = ir.CutSheet("测试-2")
	if err != nil {
		t.Log(err)
		return
	}
	// list表头
	var headers2 []*header.Header
	headers2 = append(headers2, &header.Header{Id: 101, Pid: 0, Title: "子项目名称", FieldKey: "subitem_name"})
	headers2 = append(headers2, &header.Header{Id: 102, Pid: 0, Title: "姓名", Weight: 1, FieldKey: "name"})
	headers2 = append(headers2, &header.Header{Id: 103, Pid: 0, Title: "证件号", FieldKey: "id_card"})
	headers2 = append(headers2, &header.Header{Id: 104, Pid: 0, Title: "工作城市", FieldKey: "work_city"})
	headers2 = append(headers2, &header.Header{Id: 105, Pid: 0, Title: "调整前", FieldKey: "before"})
	headers2 = append(headers2, &header.Header{Id: 1001, Pid: 105, Title: "养老", FieldKey: "before_old_age"})
	headers2 = append(headers2, &header.Header{Id: 10001, Pid: 1001, Title: "企业", FieldKey: "before_old_age_company"})
	headers2 = append(headers2, &header.Header{Id: 100001, Pid: 10001, Title: "企业基数", FieldKey: "before_old_age_company_base"})
	headers2 = append(headers2, &header.Header{Id: 100002, Pid: 10001, Title: "企业比例", Weight: 1, FieldKey: "before_old_age_company_rate"})
	headers2 = append(headers2, &header.Header{Id: 10002, Pid: 1001, Title: "个人", FieldKey: "before_old_age_personnel"})
	headers2 = append(headers2, &header.Header{Id: 100003, Pid: 10002, Title: "个人基数", FieldKey: "before_old_age_personnel_base"})
	headers2 = append(headers2, &header.Header{Id: 100004, Pid: 10002, Title: "个人比例", FieldKey: "before_old_age_personnel_rate"})
	headers2 = append(headers2, &header.Header{Id: 106, Pid: 0, Title: "调整后", FieldKey: "after"})
	headers2 = append(headers2, &header.Header{Id: 1002, Pid: 106, Title: "养老", FieldKey: "after_old_age"})
	headers2 = append(headers2, &header.Header{Id: 10003, Pid: 1002, Title: "企业", FieldKey: "after_old_age_company"})
	headers2 = append(headers2, &header.Header{Id: 100005, Pid: 10003, Title: "企业基数", FieldKey: "after_old_age_company_base"})
	headers2 = append(headers2, &header.Header{Id: 100006, Pid: 10003, Title: "企业比例", FieldKey: "after_old_age_company_rate"})
	headers2 = append(headers2, &header.Header{Id: 10004, Pid: 1002, Title: "个人", FieldKey: "after_old_age_personnel"})
	headers2 = append(headers2, &header.Header{Id: 100007, Pid: 10004, Title: "个人基数", FieldKey: "after_old_age_personnel_base"})
	headers2 = append(headers2, &header.Header{Id: 100008, Pid: 10004, Title: "个人比例", FieldKey: "after_old_age_personnel_rate"})
	headers2 = append(headers2, &header.Header{Id: 107, Pid: 0, Title: "备注", FieldKey: "remark"})
	err = ir.SetHeaderListById(headers2).Error()
	if err != nil {
		t.Log(err)
		return
	}
	dataList2 := ir.GetRows()
	for i, v := range dataList2 {
		t.Logf("------- 序号：%d ------------  %s", i, ir.PrintRows(v))
	}

	err = ir.CutSheet("测试-3")
	if err != nil {
		t.Log(err)
		return
	}
	// list表头
	var headers3 []*header.Header
	headers3 = append(headers3, &header.Header{Pkey: "", Title: "基本信息", FieldKey: "user_info"})
	headers3 = append(headers3, &header.Header{Pkey: "user_info", Title: "姓名", FieldKey: "name"})
	headers3 = append(headers3, &header.Header{Pkey: "user_info", Title: "证件号", FieldKey: "id_card"})
	headers3 = append(headers3, &header.Header{Pkey: "", Title: "子项目名称", FieldKey: "subitem_name"})
	headers3 = append(headers3, &header.Header{Pkey: "", Title: "调整前", FieldKey: "before"})
	headers3 = append(headers3, &header.Header{Pkey: "before", Title: "养老", FieldKey: "before_old_age"})
	headers3 = append(headers3, &header.Header{Pkey: "before_old_age", Title: "企业", FieldKey: "before_old_age_company"})
	headers3 = append(headers3, &header.Header{Pkey: "before_old_age_company", Title: "企业基数", FieldKey: "before_old_age_company_base"})
	headers3 = append(headers3, &header.Header{Pkey: "before_old_age_company", Title: "企业比例", Weight: 1, FieldKey: "before_old_age_company_rate"})
	headers3 = append(headers3, &header.Header{Pkey: "before_old_age", Title: "个人", FieldKey: "before_old_age_personnel"})
	headers3 = append(headers3, &header.Header{Pkey: "before_old_age_personnel", Title: "个人基数", FieldKey: "before_old_age_personnel_base"})
	headers3 = append(headers3, &header.Header{Pkey: "before_old_age_personnel", Title: "个人比例", FieldKey: "before_old_age_personnel_rate"})
	headers3 = append(headers3, &header.Header{Pkey: "", Title: "调整后", FieldKey: "after"})
	headers3 = append(headers3, &header.Header{Pkey: "after", Title: "养老", FieldKey: "after_old_age"})
	headers3 = append(headers3, &header.Header{Pkey: "after_old_age", Title: "企业", FieldKey: "after_old_age_company"})
	headers3 = append(headers3, &header.Header{Pkey: "after_old_age_company", Title: "企业基数", FieldKey: "after_old_age_company_base"})
	headers3 = append(headers3, &header.Header{Pkey: "after_old_age_company", Title: "企业比例", FieldKey: "after_old_age_company_rate"})
	headers3 = append(headers3, &header.Header{Pkey: "after_old_age", Title: "个人", FieldKey: "after_old_age_personnel"})
	headers3 = append(headers3, &header.Header{Pkey: "after_old_age_personnel", Title: "个人基数", FieldKey: "after_old_age_personnel_base"})
	headers3 = append(headers3, &header.Header{Pkey: "after_old_age_personnel", Title: "个人比例", FieldKey: "after_old_age_personnel_rate"})
	headers3 = append(headers3, &header.Header{Pkey: "", Title: "备注", FieldKey: "remark"})
	err = ir.SetHeaderListByPkey(headers3).Error()
	if err != nil {
		t.Log(err)
		return
	}
	//err = ir.SetNoHeader(0, 4).Error()
	dataList3 := ir.GetRows()
	for i, v := range dataList3 {
		t.Logf("------- 序号：%d ------------  %s", i, ir.PrintRows(v))
	}

}
