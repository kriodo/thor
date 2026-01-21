package pdf

import (
	"errors"
	"fmt"
	"github.com/jung-kurt/gofpdf"
	"github.com/jung-kurt/gofpdf/contrib/gofpdi"
)

type WritePDFByTemplateReq struct {
	Input    string     // 输入文件路径
	Output   string     // 输出文件路径
	Contents []FieldPDF // 内容
}

// WritePDFByTemplate 在现有的PDF模板里进行写入
func WritePDFByTemplate(req *WritePDFByTemplateReq) error {
	// 验证入参
	if req.Input == "" {
		return errors.New("模板文件不能为空")
	}
	if req.Output == "" {
		return errors.New("生成文件名不能为空")
	}

	pdf := gofpdf.New("P", "mm", "", "")

	// 导入器
	importer := gofpdi.NewImporter()

	// 导入原 PDF（第 1 页）
	tpl := importer.ImportPage(pdf, req.Input, 1, "/MediaBox")
	// 关键：读取原 PDF 页面尺寸（mm）
	var (
		w float64 = 240
		h float64 = 160
	)
	tplSize := importer.GetPageSizes()
	if tplSize != nil && tplSize[1] != nil && tplSize[1]["/MediaBox"] != nil {
		w = tplSize[1]["/MediaBox"]["w"]
		h = tplSize[1]["/MediaBox"]["h"]
	}

	// 用原尺寸创建页面
	pdf.AddPageFormat("P", gofpdf.SizeType{Wd: w, Ht: h})

	// 原 PDF 作为背景，完全贴合
	importer.UseImportedTemplate(pdf, tpl, 0, 0, w, h)

	// ===== 在指定位置写内容 =====
	for i, v := range req.Contents {
		if !IsLoc(v.X) || !IsLoc(v.Y) || v.FontName == "" {
			continue
		}
		if v.PageNo <= 0 {
			v.PageNo = 1
		}
		if v.FontH <= 0 {
			v.FontH = 10
		}
		if v.FontSize <= 0 {
			v.FontSize = 20
		}
		familyStr := fmt.Sprintf("ziti-%d", i)
		pdf.AddUTF8Font(familyStr, "", v.FontName)
		pdf.SetFont(familyStr, "", v.FontSize)
		// 设置位置坐标
		pdf.SetXY(w*v.X, h*v.Y) // X=50mm, Y=100mm
		// 设置字体大小和内容
		pdf.Cell(v.FontW, v.FontH, v.Val)
	}

	// 保存
	err := pdf.OutputFileAndClose(req.Output)
	if err != nil {
		return fmt.Errorf("操作PDF失败：%+v", err)
	}
	return nil
}

type FieldPDF struct {
	PageNo   int     // 页码(默认：1)
	X        float64 // 横坐标
	Y        float64 // 横坐标
	FontW    float64 // 字体-宽
	FontH    float64 // 字体-高
	Val      string  // 值
	FontName string  // 字体
	FontSize float64 // 字体大小
}

func IsLoc(v float64) bool {
	if v >= 1 || v <= 0 {
		return false
	}
	return true
}

type InvoiceInfo struct {
	Number         FieldPDF       // 发票号码
	Date           FieldPDF       // 开票日期
	AName          FieldPDF       // 购买方-名称
	ANumber        FieldPDF       // 购买方-统一社会信用代码/纳税人识别号
	BName          FieldPDF       // 销售方-名称
	BNumber        FieldPDF       // 销售方-统一社会信用代码/纳税人识别号
	DataList       []*InvoiceData // 信息列表
	TotalAmount    FieldPDF       // 合计-金额
	TotalTaxAmount FieldPDF       // 合计-税额
	CNAmount       FieldPDF       // 价税合计-金额-大写
	Amount         FieldPDF       // 价税合计-金额-小写
	Remark         FieldPDF       // 备注
	Invoicer       FieldPDF       // 开票人
}

type InvoiceData struct {
	ProjectName FieldPDF // 项目名称
	SpecType    FieldPDF // 规格型号
	Unit        FieldPDF // 单位
	Number      FieldPDF // 数量
	UnitPrice   FieldPDF // 单价
	Amount      FieldPDF // 金额
	Rate        FieldPDF // 税率/征收率
	TaxAmount   FieldPDF // 税额
}
