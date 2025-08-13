package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/excel/constant"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

// Exporter 导出
type Exporter struct {
	curSheetName string                // 当前sheet名称
	file         *excelize.File        // 文件
	sheet        map[string]*SheetInfo // sheet数据
	err          error                 // 错误
	defStyleId   int                   // 默认样式id
	dropInfo     map[string]*DropInfo  // sheet+字段 -> 数据信息
	dropSheetMap map[string]string     // 下拉数据隐藏sheet数据对应的key
	dropLen      map[string]bool       // 是否超长度验证器
}

// NewExporter new一个导出处理器
func NewExporter(sheetName string) (*Exporter, error) {
	err := tool.CheckSheetName(sheetName)
	if err != nil {
		return nil, err
	}
	file := excelize.NewFile()
	// 创建一个工作表
	index, err := file.NewSheet(sheetName)
	if err != nil {
		return nil, fmt.Errorf("导出创建sheet失败: %+v", err)
	}
	if sheetName != constant.DefaultSheetName {
		err = file.DeleteSheet(constant.DefaultSheetName)
		if err != nil {
			return nil, fmt.Errorf("导出删除默认sheet失败: %+v", err)
		}
	}
	// 设置默认sheet
	file.SetActiveSheet(index)
	er, err := NewExporterWithFile(file)
	if err != nil {
		return nil, err
	}
	return er, nil
}

// NewExporterWithFile new一个导出处理器: 自带文件
func NewExporterWithFile(file *excelize.File) (*Exporter, error) {
	sheetLen := len(file.GetSheetList())
	if sheetLen == 0 {
		return nil, fmt.Errorf("表格无sheet数据")
	}
	er := &Exporter{
		curSheetName: file.GetSheetList()[0],
		file:         file,
		sheet:        make(map[string]*SheetInfo, 10),
		err:          nil,
		defStyleId:   0,
		dropInfo:     make(map[string]*DropInfo),
		dropSheetMap: make(map[string]string),
		dropLen:      make(map[string]bool),
	}
	// 初始化默认表头样式
	er.defStyleId = er.setHeaderStyle(header.GetExportDefaultStyle())
	if er.err != nil {
		return er, er.err
	}
	// 初始化sheet
	er.initSheetInfo()
	return er, nil
}

// GetFile 获取文件file
func (er *Exporter) GetFile() *excelize.File {
	return er.file
}

// 获取错误
func (er *Exporter) Error() error {
	return er.err
}

// SaveAs 文件存储
func (er *Exporter) SaveAs(filePath string) error {
	if er.err != nil {
		return er.err
	}
	if err := er.file.SaveAs(filePath); err != nil {
		return err
	}
	return nil
}

// reversalFieldKeyMap 反转表头
//func (er *Exporter) reversalFieldKeyMap(fieldKeyMap map[string]int) map[int]string {
//	var fieldList = make(map[int]string)
//	for key, val := range fieldKeyMap {
//		fieldList[val] = key
//	}
//	return fieldList
//}

// SetValidationString 设置序号验证
//func (er *Exporter) SetValidationString(validationDropList map[string][]string) *Exporter {
//	er.validationDropList = validationDropList
//	return er
//}
