package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

// DropInfo 下拉框数据详情
type DropInfo struct {
	UniqueKey   string   // 下拉框对应的唯一标识符（长度不能超过200）
	XIndex      string   // 横坐标(字母)
	YStartIndex uint     // 纵坐标开始
	YEndIndex   uint     // 纵坐标结束
	ValueList   []string // 下拉数据(如果有历史数据可以不填写)
}

// FieldDropInfo 字段下拉框数据详情
type FieldDropInfo struct {
	UniqueKey string   // 下拉框对应的唯一标识符（长度不能超过200）
	FieldKeys []string // 字段集合
	YEndIndex uint     // 纵坐标结束
	ValueList []string // 下拉数据(如果有历史数据可以不填写)
}

// SetDrop 设置excel下拉框属性
func (er *Exporter) SetDrop(infos []*DropInfo) error {
	for _, info := range infos {
		err := er.setDrop(info)
		if err != nil {
			return err
		}
	}
	return nil
}

// SetDropByFieldKey 设置字段下拉框属性
func (er *Exporter) SetDropByFieldKey(infos []*FieldDropInfo) error {
	for _, info := range infos {
		for _, key := range info.FieldKeys {
			err := er.setDrop(&DropInfo{
				UniqueKey:   info.UniqueKey,
				XIndex:      er.GetFieldXIndex(key),
				YStartIndex: er.GetDataStartY(),
				YEndIndex:   info.YEndIndex,
				ValueList:   info.ValueList,
			})
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// setDrop 设置excel下拉框属性
func (er *Exporter) setDrop(info *DropInfo) error {
	if er.curSheet == nil {
		return fmt.Errorf("未设置sheet: %s", info.UniqueKey)
	}
	if info.XIndex == "" || info.UniqueKey == "" {
		return fmt.Errorf("%s的Drop的参数不能为空", er.curSheet.sheetName)
	}
	// 以被引用的sheet表为基准，先创建隐藏sheet页，再设置引用该sheet的所有列
	// 只能调用一次，不然会重复创建隐藏sheet
	// 判断下拉框是否需要超限，每个下拉数据只验证一次，验证过的做个标识
	// 判断是否需要新建一个隐藏的sheet
	dataUrl, exi := er.dropSheetMap[info.UniqueKey]
	if !exi {
		if len(info.ValueList) == 0 {
			return fmt.Errorf("未设置下拉框数据: %s", info.UniqueKey)
		}
	}
	isHiddenSheet, ok := er.dropLen[info.UniqueKey]
	if !ok {
		isHiddenSheet = CheckDropLen(info.ValueList)
		er.dropLen[info.UniqueKey] = isHiddenSheet
	}
	// 判断是否需要新建一个隐藏的sheet
	if !exi && isHiddenSheet {
		// 防止sheet名称过长无法创建
		dropSheetName := tool.FormatSheetName(info.UniqueKey)
		// 新建一个sheet
		_, err := er.file.NewSheet(dropSheetName)
		if err != nil {
			return fmt.Errorf("创建sheet失败: %s %+v", info.UniqueKey, err)
		}
		// 在新建的sheet页插入数据
		length := len(info.ValueList)
		for i := 0; i < length; i++ {
			cell := fmt.Sprintf("A%d", i+1)
			err = er.file.SetCellValue(dropSheetName, cell, info.ValueList[i])
			if err != nil {
				return fmt.Errorf("SetCellValue失败: %s %+v", dropSheetName, err)
			}
		}
		// 隐藏新建的sheet页
		err = er.file.SetSheetVisible(dropSheetName, false)
		if err != nil {
			return fmt.Errorf("SetSheetVisible失败: %s %+v", dropSheetName, err)
		}
		// 数据坐标串 $A$1:$B$1
		dataUrl = fmt.Sprintf("%s!$A$1:$A$%d", dropSheetName, length)
		er.dropSheetMap[info.UniqueKey] = dataUrl
		isHiddenSheet = true
	}
	// 创建一个数据验证规则，引用新建sheet页的数据
	dv := excelize.NewDataValidation(true)
	dv.SetSqref(fmt.Sprintf("%s%d:%s%d", info.XIndex, info.YStartIndex, info.XIndex, info.YEndIndex)) // 设置较大的范围
	// 判断是否创建隐藏页形式
	if isHiddenSheet {
		dv.SetSqrefDropList(dataUrl)
	} else {
		err := dv.SetDropList(info.ValueList)
		if err != nil {
			return fmt.Errorf("SetDropList失败: %s %+v", info.UniqueKey, err)
		}
	}
	// 在第一个sheet页应用数据验证规则
	err := er.file.AddDataValidation(er.curSheet.sheetName, dv)
	if err != nil {
		return fmt.Errorf("SetDropList失败: %s %+v", info.UniqueKey, err)
	}
	key := er.curSheet.sheetName + "-" + info.XIndex
	er.dropInfo[key] = &DropInfo{
		UniqueKey:   info.UniqueKey,
		XIndex:      info.XIndex,
		YStartIndex: info.YStartIndex,
		YEndIndex:   info.YEndIndex,
		ValueList:   info.ValueList,
	}
	return nil
}

// CheckDropLen 验证下拉值是否超限(超过250需要建隐藏sheet)
func CheckDropLen(slice []string) bool {
	var (
		vd  bool
		str string
	)
	for _, s := range slice {
		str += s
	}
	if tool.LenChar(str) >= 250 {
		vd = true
	}
	return vd
}
