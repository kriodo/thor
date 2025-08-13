package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

type RowData struct {
	key       string      // key
	Value     interface{} // 值
	StyleId   int         // 样式id
	ValueType             // 值类型
}

type ValueType int

const (
	ValueCellDefault ValueType = 0 //默认
	ValueCellString  ValueType = 1 //字符串
)

// SetHeaderStartX 设置表头起始列（必须SetHeader之前操作）
func (er *Exporter) SetHeaderStartX(index uint) *Exporter {
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	if len(curSheet.headerTree) > 0 {
		er.err = fmt.Errorf("SetHeaderStartX必须在设置表头之前")
		return er
	}
	curSheet.headerStartX = index
	er.sheet[er.curSheetName] = curSheet
	return er
}

// SetHeaderStartY 设置表头起始行（必须SetHeader之前操作）
func (er *Exporter) SetHeaderStartY(index uint) *Exporter {
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	if len(curSheet.headerTree) > 0 {
		er.err = fmt.Errorf("SetHeaderStartY必须在设置表头之前")
		return er
	}
	if index <= 1 {
		index = 1
	}
	curSheet.headerStartY = index
	er.sheet[er.curSheetName] = curSheet
	return er
}

// SetDataStartX 设置数据起始行（必须SetHeader之前操作）
func (er *Exporter) SetDataStartX(index uint) *Exporter {
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	if len(curSheet.headerTree) > 0 {
		er.err = fmt.Errorf("表头数据为空，请先设置表头参数")
		return er
	}
	if index <= 1 {
		index = 1
	}
	er.sheet[er.curSheetName] = curSheet
	return er
}

// SetDataByRow 设置顺序数据 [** 此模式数据需要按照顺序赋值，优点：快|缺点：顺序必须对齐 **]
func (er *Exporter) SetDataByRow(rows [][]interface{}) *Exporter {
	if er.err != nil {
		return er
	}
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	err = er.handleDataRow(rows, curSheet.GetDataStartX())
	if err != nil {
		er.err = err
		return er
	}
	return er
}

// SetDataByMap 设置map数据 [此模式只需要对应的key附上值即可]
func (er *Exporter) SetDataByMap(dataMap []map[string]interface{}) *Exporter {
	if er.err != nil {
		return er
	}
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	rows := make([][]interface{}, 0, len(dataMap))
	for _, data := range dataMap {
		row := make([]interface{}, curSheet.headerFieldLen)
		for key, val := range data {
			if info, exi := curSheet.fieldInfoMap[key]; exi {
				if info.XIndex > curSheet.headerFieldLen {
					er.err = fmt.Errorf("字段的索引错误: %s-%d", key, info.XIndex)
					return er
				}
				row[info.XIndex] = val
			}
		}
		rows = append(rows, row)
	}
	er.SetDataByRow(rows)
	return er
}

// SetDataMapWithStyle 设置数据 + 设置样式 [此模式只需要对应的key附上值即可]
func (er *Exporter) SetDataMapWithStyle(dataMap []map[string]*RowData) *Exporter {
	if er.err != nil {
		return er
	}
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		er.err = err
		return er
	}
	rows := make([][]*RowData, 0, len(dataMap))
	for _, data := range dataMap {
		row := make([]*RowData, curSheet.headerFieldLen)
		for key, val := range data {
			if info, exi := curSheet.fieldInfoMap[key]; exi {
				if info.XIndex > curSheet.headerFieldLen {
					er.err = fmt.Errorf("字段的索引错误: %s-%d", key, info.XIndex)
					return er
				}
				val.key = key
				row[info.XIndex] = val
			}
		}
		rows = append(rows, row)
	}
	err = er.handleData(rows, curSheet.GetDataStartX())
	if err != nil {
		er.err = err
		return er
	}
	return er
}

func (er *Exporter) handleDataRow(rows [][]interface{}, dataStartLine uint) error {
	for i, row := range rows {
		index := i + int(dataStartLine)
		letter := tool.IndexToLetter(uint(i))
		rowAddr, err := excelize.JoinCellName(letter, index)
		if err != nil {
			return fmt.Errorf("JoinCellName失败: %s-%d-%+v", letter, index, err)
		}
		if err = er.file.SetSheetRow(er.curSheetName, rowAddr, &row); err != nil {
			return fmt.Errorf("SetSheetRow失败: %+v", err)
		}
	}
	return nil
}

func (er *Exporter) handleData(rows [][]*RowData, dataStartLine uint) error {
	curSheet, err := er.GetSheetInfo(er.curSheetName)
	if err != nil {
		return err
	}
	for i, row := range rows {
		for ii, info := range row {
			var (
				columStyleId = curSheet.columnStyle[info.key] // 列的style
				index        = i + int(dataStartLine)         // 当前字段索引
				letter       = tool.IndexToLetter(uint(ii))   // 当前位置字母
			)
			rowAddr, subErr := excelize.JoinCellName(letter, index)
			switch info.ValueType {
			case ValueCellString: // 字符串
				subErr = er.file.SetCellStr(er.curSheetName, rowAddr, tool.Any2String(info.Value))
				if subErr != nil {
					return fmt.Errorf("SetCellStr失败: %s-%d-%+v", letter, index, subErr)
				}
			default:
				subErr = er.file.SetCellValue(er.curSheetName, rowAddr, info.Value)
				if subErr != nil {
					return fmt.Errorf("SetCellValue失败: %s-%d-%+v", letter, index, subErr)
				}
			}
			// 优先列样式，字段单独样式次之
			if columStyleId <= 0 && info.StyleId > 0 {
				if subErr = er.file.SetCellStyle(er.curSheetName, rowAddr, rowAddr, info.StyleId); subErr != nil {
					return fmt.Errorf("SetCellStyle失败: %s-%d-%+v", letter, index, subErr)
				}
			}
		}
	}

	return nil
}
