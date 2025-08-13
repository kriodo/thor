package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/kriodo/thor/office/tool"
)

// SheetInfo sheet相关数据
type SheetInfo struct {
	sheetName      string                       // 表头tree数据
	headerTree     []*header.Header             // 表头tree数据
	fieldInfoList  []*header.FieldInfo          // 字段数据 list
	headerFieldLen uint                         // 字段个数
	fieldInfoMap   map[string]*header.FieldInfo // 字段数据 map
	styleId        int                          // 样式id: 默认同全局
	headerStartX   uint                         // 表头起始行号: 默认=0
	headerStartY   uint                         // 表头起始列号: 默认=1
	headerMaxLevel uint                         // 表头最大层数: 默认=1
	dataStartY     uint                         // 数据起始列号: 默认表头结束行号+1
	columnStyle    map[string]int               // 列的单元格样式
}

// GetSheetInfo 获取sheet数据
func (er *Exporter) GetSheetInfo(sheetName string) (*SheetInfo, error) {
	info, exi := er.sheet[sheetName]
	if !exi {
		return nil, fmt.Errorf(sheetName + "无数据")
	}
	return info, nil
}

// AddSheet 增加一个sheet
func (er *Exporter) AddSheet(sheetName string) (*Exporter, error) {
	if er.err != nil {
		return er, nil
	}
	err := tool.CheckSheetName(sheetName)
	if err != nil {
		return nil, err
	}
	if tool.InStringArray(sheetName, er.file.GetSheetList()) {
		return nil, fmt.Errorf(sheetName + "已存在，无法创建")
	}
	_, err = er.file.NewSheet(sheetName)
	if err != nil {
		return er, fmt.Errorf("创建sheet失败: %s %+v", sheetName, err)
	}
	er.curSheetName = sheetName
	er.initSheetInfo()
	return er, nil
}

// CutSheet 切换sheet
func (er *Exporter) CutSheet(sheetName string) *Exporter {
	er.curSheetName = sheetName
	return er
}

// GetFieldXIndex 获取字段的x位置字母坐标
func (er *Exporter) GetFieldXIndex(filedKey string) string {
	if er.err != nil {
		return ""
	}
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return ""
	}
	val, ok := curSheet.fieldInfoMap[filedKey]
	if !ok {
		return ""
	}
	return tool.IndexToLetter(val.XIndex)
}

// GetDataStartY 获取数据的Y起始坐标
func (er *Exporter) GetDataStartY(filedKey string) uint {
	if er.err != nil {
		return 0
	}
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return 0
	}
	return curSheet.getDataStartY()
}

// GetHeaderEndY 获取表头结束行号y值
func (si *SheetInfo) GetHeaderEndY() uint {
	if si.headerStartY > 1 {
		return si.headerMaxLevel + si.headerStartY - 1
	}
	return si.headerMaxLevel
}

// GetDataStartX 获取数据开始列号x值
func (si *SheetInfo) GetDataStartX() uint {
	return si.headerStartX + 1
}

// 获取数据开始行号y值
func (si *SheetInfo) getDataStartY() uint {
	return si.GetHeaderEndY() + 1
}

// 初始化一个sheet
func (er *Exporter) initSheetInfo() {
	er.sheet[er.curSheetName] = &SheetInfo{
		sheetName:      er.curSheetName,
		headerTree:     nil,
		fieldInfoList:  nil,
		fieldInfoMap:   make(map[string]*header.FieldInfo),
		styleId:        er.defStyleId,
		headerStartX:   0,
		headerStartY:   1,
		headerMaxLevel: 1,
		headerFieldLen: 0,
		columnStyle:    nil,
	}
}
