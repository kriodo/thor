package pdf

import (
	"os"
	"testing"
)

func TestAddNewPDF(t *testing.T) {
	os.Remove("output.pdf")
	//var (
	//	w float64 = 240
	//	h float64 = 160
	//)
	var contents []FieldPDF
	contents = append(contents, FieldPDF{
		X:        0.81,
		Y:        0.08,
		Val:      "24112000000050652269",
		FontName: "华文细黑.ttf",
		FontSize: 27,
	})
	contents = append(contents, FieldPDF{
		X:   0.83,
		Y:   0.125,
		Val: "2024年05月30日",
	})
	contents = append(contents, FieldPDF{
		X:        0.09,
		Y:        0.24,
		Val:      "北京三快在线科技有限公司",
		FontName: "兰米仿宋.ttf",
		FontSize: 30,
	})
	contents = append(contents, FieldPDF{
		X:   0.27,
		Y:   0.31,
		Val: "91110108562144110X",
	})
	contents = append(contents, FieldPDF{
		X:   0.55,
		Y:   0.24,
		Val: "北京万古恒信科技有限公司",
	})

	contents = append(contents, FieldPDF{
		X:   0.73,
		Y:   0.31,
		Val: "91110102673821137F",
	})
	//---------- 多行数据部分------------//
	contents = append(contents, FieldPDF{
		X:   0.02,
		Y:   0.403,
		Val: "*现代服务*技术服务",
	})
	contents = append(contents, FieldPDF{
		X:   0.67,
		Y:   0.403,
		Val: "13893.76",
	})
	contents = append(contents, FieldPDF{
		X:   0.78,
		Y:   0.403,
		Val: "6%",
	})
	contents = append(contents, FieldPDF{
		X:   0.92,
		Y:   0.403,
		Val: "833.63",
	})

	contents = append(contents, FieldPDF{
		X:   0.02,
		Y:   0.433,
		Val: "*现代服务*技术服务",
	})
	contents = append(contents, FieldPDF{
		X:   0.67,
		Y:   0.433,
		Val: "13893.76",
	})
	contents = append(contents, FieldPDF{
		X:   0.78,
		Y:   0.433,
		Val: "6%",
	})
	contents = append(contents, FieldPDF{
		X:   0.92,
		Y:   0.433,
		Val: "833.63",
	})
	//---------- 多行数据部分------------//

	contents = append(contents, FieldPDF{
		X:   0.65,
		Y:   0.653,
		Val: "¥13893.76",
	})
	contents = append(contents, FieldPDF{
		X:   0.92,
		Y:   0.653,
		Val: "¥833.63",
	})
	contents = append(contents, FieldPDF{
		X:   0.30,
		Y:   0.70,
		Val: "壹万肆仟柒佰贰拾柒圆叁角玖分",
	})
	contents = append(contents, FieldPDF{
		X:   0.73,
		Y:   0.70,
		Val: "¥14727.39",
	})
	contents = append(contents, FieldPDF{
		X:   0.054,
		Y:   0.743,
		Val: "PO2405111108274526 这是一个备注",
	})
	contents = append(contents, FieldPDF{
		X:   0.15,
		Y:   0.924395,
		Val: "杨灿",
	})

	err := WritePDFByTemplate(&WritePDFByTemplateReq{
		Input:    "tpl-invoice.pdf",
		Output:   "output.pdf",
		Contents: contents,
	})
	if err != nil {
		t.Log(err)
		return
	}
}
