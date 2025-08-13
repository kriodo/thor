package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

// SetTree 设置tree数据表头 [** 此模式id|pid|pkey无需填写 **]
func (er *Exporter) SetTree(tree []*header.Header) *Exporter {
	if er.err != nil {
		return er
	}
	curSheet, err := er.GetCurSheetInfo()
	if err != nil {
		er.err = err
		return er
	}
	// 获取表头相关信息
	curSheet.headerTree = tree
	curSheet.fieldInfoList, er.err = header.FormatHeaderInfo(curSheet.headerTree, 1, curSheet.fieldInfoList)
	if er.err != nil {
		return er
	}
	for i, v := range curSheet.fieldInfoList {
		index := uint(i) + curSheet.headerStartX
		if v.YIndex > curSheet.headerMaxLevel {
			curSheet.headerMaxLevel = v.YIndex
		}
		v.XIndex = index
		curSheet.fieldInfoMap[v.Key] = v
	}
	// 处理表头
	er.err = er.handleHeader(curSheet.headerTree, curSheet.headerStartX, curSheet.headerStartY)
	if er.err != nil {
		return er
	}
	// 设置最后表头数据宽度
	er.err = er.setHeaderWidth()
	if er.err != nil {
		return er
	}
	er.sheet[er.curSheet.sheetName] = curSheet
	er.curSheet = curSheet
	return er
}

// SetListById 设置list数据 [此模式id、pid必须填写]
func (er *Exporter) SetListById(headers []*header.Header) *Exporter {
	if er.err != nil {
		return er
	}
	err := header.CheckHeaderId(headers)
	if err != nil {
		er.err = err
		return er
	}
	// 将list数据转为tree数据
	tree := header.ListToTree(headers, 0)
	er.SetTree(tree)
	return er
}

// SetListByPkey 设置list数据 [此模式pkey必须填写]
func (er *Exporter) SetListByPkey(headers []*header.Header) *Exporter {
	var (
		newHeaders  []*header.Header
		headerIdMap = make(map[string]int64)
		id          int64
	)
	for _, head := range headers {
		if head.Pkey == "" {
			er.err = fmt.Errorf("PKEY不能为空: " + head.FieldKey)
			return er
		}
		id++
		head.Id = id
		headerIdMap[head.FieldKey] = id
	}
	for _, head := range headers {
		head.Pid = headerIdMap[head.Pkey]
		newHeaders = append(newHeaders, head)
	}
	return er.SetListById(newHeaders)
}

// 设置多级表头，会根据层级自动合并
func (er *Exporter) handleHeader(headers []*header.Header, x, y uint) error {
	if er.err != nil {
		return er.err
	}
	curSheet, err := er.GetCurSheetInfo()
	if err != nil {
		return err
	}
	for _, v := range headers {
		err = er.left2right(v, x, y)
		if err != nil {
			return err
		}
		if v.GetIsLast() {
			if curSheet.GetHeaderEndY() > v.GetLevel() {
				err = er.up2down(v, x, y)
				if err != nil {
					return err
				}
			}
		} else {
			err = er.handleHeader(v.Children, x, y+1)
			if err != nil {
				return err
			}
		}
		x += header.ChildrenLen(v.Children)
	}
	return nil
}

// 左右合并
func (er *Exporter) left2right(val *header.Header, x uint, y uint) error {
	if er.err != nil {
		return er.err
	}
	curSheet, err := er.GetCurSheetInfo()
	if err != nil {
		return err
	}
	leftIndex := x
	// 左右合并单元格
	leftLetter := tool.ExcelIndexLetter(leftIndex)
	left, err := excelize.JoinCellName(leftLetter, int(y))
	if err != nil {
		return fmt.Errorf("JoinCellNameLeft失败: %s-%d %+v", leftLetter, y, err)
	}
	if err = er.file.SetCellValue(curSheet.sheetName, left, val.Title); err != nil {
		return fmt.Errorf("SetCellValue失败: %s-%s-%s %+v", curSheet.sheetName, left, val.Title, err)
	}
	// 行合并单元格
	rightIndex := leftIndex + header.ChildrenLen(val.Children) - 1
	rightLetter := tool.ExcelIndexLetter(rightIndex)
	right, err := excelize.JoinCellName(rightLetter, int(y))
	if err != nil {
		return fmt.Errorf("JoinCellNameRight失败: %s-%d %+v", rightLetter, y, err)
	}
	if err = er.file.MergeCell(curSheet.sheetName, left, right); err != nil {
		return fmt.Errorf("MergeCell失败: %s-%s-%s %+v", curSheet.sheetName, left, right, err)
	}
	// 为标题行设置样式
	styleId := curSheet.styleId
	if val.Export.StyleId > 0 {
		styleId = val.Export.StyleId
	}
	if err = er.file.SetCellStyle(curSheet.sheetName, left, right, styleId); err != nil {
		return fmt.Errorf("SetCellStyle失败: %s-%s-%s-%d %+v", curSheet.sheetName, left, right, styleId, err)
	}
	return nil
}

// 上下合并
func (er *Exporter) up2down(val *header.Header, x uint, y uint) error {
	if er.err != nil {
		return er.err
	}
	curSheet, err := er.GetCurSheetInfo()
	if err != nil {
		return err
	}
	letter := tool.ExcelIndexLetter(x)
	up, err := excelize.JoinCellName(letter, int(y))
	if err != nil {
		return fmt.Errorf("JoinCellNameUp失败: %s-%d %+v", letter, y, err)
	}
	down, err := excelize.JoinCellName(letter, int(curSheet.GetHeaderEndY()))
	if err != nil {
		return fmt.Errorf("JoinCellNameDown失败: %s-%d %+v", letter, curSheet.GetHeaderEndY(), err)
	}
	if err = er.file.SetCellValue(curSheet.sheetName, up, val.Title); err != nil {
		return fmt.Errorf("SetCellValue失败: %s-%s-%s %+v", curSheet.sheetName, up, val.Title, err)
	}
	// 合并单元格
	if err = er.file.MergeCell(curSheet.sheetName, up, down); err != nil {
		return fmt.Errorf("MergeCell失败: %s-%s-%s %+v", curSheet.sheetName, up, down, err)
	}
	// 为标题行设置样式
	styleId := curSheet.styleId
	if val.Export.StyleId > 0 {
		styleId = val.Export.StyleId
	}
	if err = er.file.SetCellStyle(curSheet.sheetName, up, down, styleId); err != nil {
		return fmt.Errorf("SetCellStyle失败: %s-%s-%s-%d %+v", curSheet.sheetName, up, down, styleId, err)
	}
	// 批注
	if val.Export.Comment != "" {
		if err = er.file.AddComment(curSheet.sheetName, excelize.Comment{
			Cell: fmt.Sprintf("%s%d", letter, y),
			Paragraph: []excelize.RichTextRun{
				{Text: val.Export.Comment},
			},
		}); err != nil {
			return fmt.Errorf("AddComment失败: %s %+v", letter, err)
		}
	}
	return nil
}
