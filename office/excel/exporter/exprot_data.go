package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

type Data struct {
	Val     interface{} // 值
	valType ValueType   // 值类型（*** 建议按照列进行赋值的类型 ***）
	styleId int         // 样式id
}

type DataOption func(*Data)

func GetData(opts ...DataOption) *Data {
	// 默认值
	opt := &Data{}
	// 应用调用者传入的配置
	for _, o := range opts {
		o(opt)
	}
	return opt
}

func SetVal(val interface{}) DataOption {
	return func(o *Data) {
		o.Val = val
	}
}

func SetValType(vt ValueType) DataOption {
	return func(o *Data) {
		o.valType = vt
	}
}

func SetStyleId(styleId int) DataOption {
	return func(o *Data) {
		o.styleId = styleId
	}
}

type ValueType int

const (
	DEFAULT ValueType = 0  // 默认
	STRING  ValueType = 10 // 字符串
	NUMBER  ValueType = 20 // 数字
	INT     ValueType = 21 // 整数
	FLOAT   ValueType = 22 // 浮点
)

// SetDataStartX 设置数据起始行（必须SetHeader之前操作）
func (er *Exporter) SetDataStartX(index uint) *Exporter {
	curSheet, err := er.GetCurSheetInfo()
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
	er.sheet[er.cur.sheetName] = curSheet
	return er
}

// SetDataBySlice 设置顺序数据 [** 此模式数据需要按照顺序赋值，优点：快|缺点：顺序必须对齐 **]
func (er *Exporter) SetDataBySlice(rows [][]*Data) *Exporter {
	if er.err != nil {
		return er
	}
	err := er.handleData(rows)
	if err != nil {
		er.err = err
		return er
	}
	return er
}

// SetDataByMap 设置数据  [此模式只需要对应的key附上值即可]
func (er *Exporter) SetDataByMap(dataMap []map[string]*Data) *Exporter {
	if er.err != nil {
		return er
	}
	curSheet, err := er.GetCurSheetInfo()
	if err != nil {
		er.err = err
		return er
	}
	rows := make([][]*Data, 0, len(dataMap))
	headerFieldLen := len(curSheet.fields)
	for _, data := range dataMap {
		row := make([]*Data, headerFieldLen)
		for key, val := range data {
			if info, exi := curSheet.fieldMap[key]; exi {
				if int(info.XIndex) > headerFieldLen {
					er.err = fmt.Errorf("字段的索引错误: %s-%d", key, info.XIndex)
					return er
				}
				row[info.XIndex] = val
			}
		}
		rows = append(rows, row)
	}
	err = er.handleData(rows)
	if err != nil {
		er.err = err
		return er
	}
	return er
}

// ----------------------------------------[ 分割线 ]-------------------------------------------//

func (er *Exporter) handleData(rows [][]*Data) error {
	if er.cur.dataLen > 0 {
		return fmt.Errorf("禁止重复设置数据: %s", er.cur.sheetName)
	}
	startX := er.cur.GetDataStartX()
	startY := er.cur.GetDataStartY()
	for i, row := range rows {
		y := i + int(startY)
		for ii, v := range row {
			x := ii + int(startX) // 当前字段索引
			err := er.handleDataOne(v, x, y)
			if err != nil {
				return err
			}
		}
		er.cur.dataLen++
	}

	return nil
}

func (er *Exporter) handleDataOne(info *Data, x int, y int) error {
	if info == nil {
		return nil
	}
	var (
		xIndex       = tool.IndexToLetter(uint(x)) // 当前字段位置字母
		columStyleId = er.cur.xStyle[xIndex]       // 列的style
	)
	rowAddr, err := excelize.JoinCellName(xIndex, y)
	if err != nil {
		return fmt.Errorf("JoinCellName失败: %s-%d-%+v", xIndex, x, err)
	}
	// 处理数据类型
	switch info.valType {
	case STRING: // 字符串
		err = er.file.SetCellStr(er.cur.sheetName, rowAddr, tool.Any2String(info.Val))
		if err != nil {
			return fmt.Errorf("SetCellStr失败: %s-%d-%+v", xIndex, x, err)
		}
	default:
		err = er.file.SetCellValue(er.cur.sheetName, rowAddr, info.Val)
		if err != nil {
			return fmt.Errorf("SetCellValue失败: %s-%d-%+v", xIndex, x, err)
		}
	}
	// 值样式：优先列样式，字段单独样式次之
	if columStyleId <= 0 && info.styleId > 0 {
		if err = er.file.SetCellStyle(er.cur.sheetName, rowAddr, rowAddr, info.styleId); err != nil {
			return fmt.Errorf("SetCellStyle失败: %s-%d-%+v", xIndex, x, err)
		}
	}
	return nil
}
