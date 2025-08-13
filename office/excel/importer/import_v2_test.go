package importer

import (
	"encoding/json"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/xuri/excelize/v2"
	"os"
	"strings"
	"testing"
)

func TestNewImportProcessor(t *testing.T) {
	f, err := excelize.OpenFile("./file/导入demo2.xlsx")
	if err != nil {
		t.Log(err)
		return
	}
	defer f.Close()
	ipv2 := NewImportProcessor(f)
	var headers []*header.Header
	//headers = []*Header{
	//	{
	//		Title:    "子项目名称",
	//		FieldKey: "suitem_name",
	//		Children: nil,
	//	},
	//	{
	//		Title:    "基础信息",
	//		FieldKey: "before",
	//		Children: []*Header{
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
	//err = ipv2.SetTree(headers).Run()

	//err = json.Unmarshal([]byte(headerStr), &headers)
	//if err != nil {
	//	return
	//}
	headers = append(headers, header.Header{Id: 1, Pid: 0, Title: "子项目名称", FieldKey: "suitem_name", MustExi: true})
	headers = append(headers, header.Header{Id: 2, Pid: 0, Title: "养老", FieldKey: "old", MustExi: true})
	headers = append(headers, header.Header{Id: 3, Pid: 2, Title: "企业", FieldKey: "old_company", MustExi: true})
	headers = append(headers, header.Header{Id: 4, Pid: 2, Title: "个人", FieldKey: "old_person", MustExi: true})
	headers = append(headers, header.Header{Id: 5, Pid: 3, Title: "基数", FieldKey: "old_company_base", MustExi: true})
	headers = append(headers, header.Header{Id: 6, Pid: 4, Title: "基数", FieldKey: "old_person_base", MustExi: true})
	err = ipv2.SetList(headers).SetStartLine(1, 5).Run()
	if err != nil {
		t.Log(err)
		return
	}
	for i, m := range ipv2.GetRows() {
		for key, val := range m {
			t.Logf("------> 序号 %d %s    %s", i+1, key, val)
		}
	}
	t.Log("--------success----------")
}

func TestNewImportProcessor222(t *testing.T) {
	f, err := excelize.OpenFile("./file/222.xlsx")
	if err != nil {
		t.Log(err)
		return
	}
	defer f.Close()
	ipv2 := NewImportProcessor(f)
	var headers []*header.Header
	js, err := os.ReadFile("./file/222.json")
	if err != nil {
		t.Log(err)
		return
	}
	if err = json.Unmarshal(js, &headers); err != nil {
		t.Log(err)
		return
	}

	err = ipv2.SetTree(headers).SetOpts(excelize.Options{
		RawCellValue: false,
	}).Run()
	//err = ipv2.SetNoHeader(1, 3).Run()
	if err != nil {
		t.Log(err)
		return
	}
	for i, m := range ipv2.GetRows() {
		for key, val := range m {
			if !strings.HasSuffix(key, "_rate") {
				continue
			}
			t.Logf("------> 序号 %d %s    %s", i+1, key, val)
		}
	}
	t.Log("--------success----------")
}
