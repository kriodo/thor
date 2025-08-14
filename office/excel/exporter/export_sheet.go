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
	fieldInfos     []*header.FieldInfo          // 字段数据 list
	headerFieldLen int                          // 字段个数
	fieldInfoMap   map[string]*header.FieldInfo // 字段数据 map
	styleId        int                          // 样式id: 默认同全局
	headerStartX   uint                         // 表头起始行号: 默认=0
	headerStartY   uint                         // 表头起始列号: 默认=1
	headerMaxLevel uint                         // 表头最大层数: 默认=1
	dataStartY     uint                         // 数据起始列号(在表头Y的基础上相加)
	xStyle         map[string]int               // 列的单元格样式: 字母 -> SetStyleId
	dataLen        uint                         // 数据量
}

// GetCurSheetInfo 获取当前sheet数据
func (er *Exporter) GetCurSheetInfo() (*SheetInfo, error) {
	if er.cur == nil || er.cur.sheetName == "" {
		return nil, fmt.Errorf("当前sheet无数据")
	}
	return er.cur, nil
}

// AddSheet 增加一个sheet
func (er *Exporter) AddSheet(sheetName string) (*Exporter, error) {
	if er.err != nil {
		return er, nil
	}
	if tool.InStringArray(sheetName, er.file.GetSheetList()) {
		return nil, fmt.Errorf(sheetName + "已存在，无法创建")
	}
	err := tool.CheckSheetName(sheetName)
	if err != nil {
		return nil, err
	}
	_, err = er.file.NewSheet(sheetName)
	if err != nil {
		return er, fmt.Errorf("创建sheet失败: %s %+v", sheetName, err)
	}
	er.initSheetInfo(sheetName)
	return er, nil
}

// CutSheet 切换sheet
func (er *Exporter) CutSheet(sheetName string) *Exporter {
	if _, exi := er.sheet[sheetName]; !exi {
		er.err = fmt.Errorf("%s不存在", sheetName)
	}
	er.cur = er.sheet[sheetName]
	return er
}

// GetFieldXIndex 获取字段的x位置字母坐标
func (er *Exporter) GetFieldXIndex(filedKey string) string {
	if er.err != nil {
		return ""
	}
	curSheet, err := er.GetCurSheetInfo()
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
func (er *Exporter) GetDataStartY() uint {
	if er.err != nil {
		return 0
	}
	curSheet, err := er.GetCurSheetInfo()
	if err != nil {
		er.err = err
		return 0
	}
	return curSheet.GetDataStartY()
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
	return si.headerStartX
}

// GetDataStartY 获取数据开始行号y值
func (si *SheetInfo) GetDataStartY() uint {
	return si.GetHeaderEndY() + si.dataStartY + 1
}

// 初始化一个sheet
func (er *Exporter) initSheetInfo(sheetName string) {
	sheet := &SheetInfo{
		sheetName:      sheetName,
		headerTree:     nil,
		fieldInfos:     nil,
		headerFieldLen: 0,
		fieldInfoMap:   make(map[string]*header.FieldInfo),
		styleId:        er.defStyleId,
		headerStartX:   0,
		headerStartY:   1,
		headerMaxLevel: 1,
		dataStartY:     0,
		xStyle:         make(map[string]int),
		dataLen:        0,
	}
	er.cur = sheet
	er.sheet[sheetName] = sheet
}
