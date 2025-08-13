package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

// SetStyle 设置样式
func (er *Exporter) SetStyle(headerStyle *header.Style) *Exporter {
	if headerStyle == nil {
		return er
	}
	sheetInfo, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	// 定义边框样式
	border := []excelize.Border{
		{Type: "top", Style: 1, Color: headerStyle.BorderColor},
		{Type: "left", Style: 1, Color: headerStyle.BorderColor},
		{Type: "right", Style: 1, Color: headerStyle.BorderColor},
		{Type: "bottom", Style: 1, Color: headerStyle.BorderColor},
	}
	// 定义标题行单元格样式
	sheetInfo.styleId, err = er.file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:   false,
			Family: headerStyle.FontFamily,
			Size:   headerStyle.FontSize,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{headerStyle.FillColor},
			Pattern: 1,
		},
		Border:       border,
		CustomNumFmt: headerStyle.CustomNumFmt,
		Alignment: &excelize.Alignment{
			Vertical:   "center",
			Horizontal: "center",
		}},
	)
	if err != nil {
		er.err = fmt.Errorf("设置样式失败: %+v", err)
		return er
	}
	return er
}

// SetColStyle 设置列的单元格样式 [ 字段:StyleId ]
func (er *Exporter) SetColStyle(columnStyle map[string]int) *Exporter {
	if er.err != nil {
		return er
	}
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	curSheet.columnStyle = columnStyle
	return er
}

// SetStringStyle 设置列的文本格式
// FieldKeys 字段key
// StartLine 开始行号: <=0时候默认为1
// EndLine   结束行号: <=0时候默认为100
func (er *Exporter) SetStringStyle(fieldKeys []string, startLine, endLine uint) error {
	if er.err != nil {
		return er.err
	}
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		return err
	}
	if len(fieldKeys) == 0 {
		return nil
	}
	if startLine <= 0 {
		startLine = 1
	}
	if endLine <= 0 {
		startLine = 100
	}
	fieldKeys = tool.UniqueString(fieldKeys)
	// 创建一个新的样式
	strStyleId, err := er.GetFile().NewStyle(&excelize.Style{NumFmt: 49}) // 49="@" 表示文本格式
	if err != nil {
		return err
	}
	for _, key := range fieldKeys {
		if val, exi := curSheet.fieldInfoMap[key]; exi {
			keyAbc := tool.IndexToLetter(val.XIndex)
			err = er.GetFile().SetCellStyle(er.curSheetName, fmt.Sprintf("%s%d", keyAbc, startLine), fmt.Sprintf("%s%d", keyAbc, endLine), strStyleId)
			if err != nil {
				return fmt.Errorf("SetCellStyle失败: %s %+v", keyAbc, err)
			}
		}
	}
	return nil
}

// 根据表头进行设置表头合适的宽度
func (er *Exporter) setHeaderWidth(tree []*header.Header, xIndex uint) (uint, error) {
	var err error
	for _, head := range tree {
		if len(head.Children) == 0 {
			x := tool.IndexToLetter(xIndex)
			// TODO 可以优化
			w := float64(tool.LenChar(head.Title))*1.7 + 8
			err = er.file.SetColWidth(er.curSheetName, x, x, w)
			if err != nil {
				return 0, err
			}
			xIndex = xIndex + 1
		} else {
			return er.setHeaderWidth(head.Children, xIndex)
		}
	}
	return xIndex, nil
}

// 设置表头样式
func (er *Exporter) setHeaderStyle(headerStyle *header.Style) int {
	if headerStyle == nil {
		return 0
	}
	// 定义边框样式
	border := []excelize.Border{
		{Type: "top", Style: 1, Color: headerStyle.BorderColor},
		{Type: "left", Style: 1, Color: headerStyle.BorderColor},
		{Type: "right", Style: 1, Color: headerStyle.BorderColor},
		{Type: "bottom", Style: 1, Color: headerStyle.BorderColor},
	}
	// 定义标题行单元格样式
	styleId, err := er.file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:   false,
			Family: headerStyle.FontFamily,
			Size:   headerStyle.FontSize,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{headerStyle.FillColor},
			Pattern: 1,
		},
		Border:       border,
		CustomNumFmt: headerStyle.CustomNumFmt,
		Alignment: &excelize.Alignment{
			Vertical:   "center",
			Horizontal: "center",
		}},
	)
	if err != nil {
		er.err = fmt.Errorf("设置样式失败: %+v", err)
		return 0
	}
	return styleId
}
