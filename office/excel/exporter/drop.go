package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

type DropListV2Req struct {
	DropList    []string // 下拉数据
	DropKey     string   // 下拉框对应的唯一标识符
	XIndex      int      // 横坐标
	YStartIndex int      // 纵坐标开始
	YEndIndex   int      // 纵坐标结束
}

type QuoteDropXY struct {
	XIndex      string // 横坐标
	YStartIndex int    // 纵坐标开始
	YEndIndex   int    // 纵坐标结束
	DropKey     string // 下拉数据坐标
}

type DropResp struct {
	f             *excelize.File          // excel
	sheetName     string                  // 数据sheet
	dataMap       map[string]*QuoteDropXY // x => 数据信息
	quoteSheetMap map[string]string       // 下拉数据隐藏sheet数据对应的key
	verifyMap     map[string]int          // 是否超限验证器: 1=超限 2=不超限
}

// NewDrop 初始化一个excel下拉设置
func (er *Exporter) NewDrop(sheetName string) *DropResp {
	return &DropResp{
		f:             er.file,
		sheetName:     sheetName,
		dataMap:       make(map[string]*QuoteDropXY),
		quoteSheetMap: make(map[string]string),
		verifyMap:     make(map[string]int),
	}
}

// SetDropListV2 批量设置excel下拉框属性
func (d *DropResp) SetDropListV2(req *DropListV2Req) error {
	if len(req.DropList) == 0 {
		return nil
	}
	// 以被引用的sheet表为基准，先创建隐藏sheet页，再设置引用该sheet的所有列
	// 只能调用一次，不然会重复创建隐藏sheet
	var (
		isHiddenSheet bool
		err           error
	)
	// 判断下拉框是否需要超限，每个下拉数据只验证一次，验证过的做个标识
	vd, ok := d.verifyMap[req.DropKey]
	if !ok {
		vd = VerifyDrop(req.DropList)
	}
	isHiddenSheet = vd == 1
	d.verifyMap[req.DropKey] = vd

	// 判断是否需要新建一个隐藏的sheet
	dataUrl, exi := d.quoteSheetMap[req.DropKey]
	if !exi && d.verifyMap[req.DropKey] == 1 {
		// 防止sheet名称过长无法创建
		dropSheetName := tool.FormatSheetName(req.DropKey)
		// 新建一个sheet
		_, err = d.f.NewSheet(dropSheetName)
		if err != nil {
			return fmt.Errorf("excel drop failed. %s:%+v", req.DropKey, err)
		}
		// 在新建的sheet页插入数据
		length := len(req.DropList)
		for i := 0; i < length; i++ {
			cell := fmt.Sprintf("A%d", i+1)
			err = d.f.SetCellValue(dropSheetName, cell, req.DropList[i])
			if err != nil {
				return fmt.Errorf("excel drop failed. %s:%+v", req.DropKey, err)
			}
		}
		// 隐藏新建的sheet页
		err = d.f.SetSheetVisible(dropSheetName, false)
		if err != nil {
			return fmt.Errorf("excel drop failed. %s:%+v", req.DropKey, err)
		}
		// 数据坐标串 $A$1:$B$1
		dataUrl = fmt.Sprintf("%s!$A$1:$A$%d", dropSheetName, length)
		d.quoteSheetMap[req.DropKey] = dataUrl
		isHiddenSheet = true
	}

	// 创建一个数据验证规则，引用新建sheet页的数据
	dv := excelize.NewDataValidation(true)
	x := tool.IndexToLetter(req.XIndex)
	dv.SetSqref(fmt.Sprintf("%s%d:%s%d", x, req.YStartIndex, x, req.YEndIndex)) // 设置较大的范围
	// 判断是否创建隐藏页形式
	if isHiddenSheet {
		dv.SetSqrefDropList(dataUrl)
	} else {
		err = dv.SetDropList(req.DropList)
		if err != nil {
			return fmt.Errorf("excel drop failed. %s:%+v", req.DropKey, err)
		}
	}
	// 在第一个sheet页应用数据验证规则
	err = d.f.AddDataValidation(d.sheetName, dv)
	if err != nil {
		return fmt.Errorf("excel drop failed. %s:%+v", req.DropKey, err)
	}
	d.dataMap[x] = &QuoteDropXY{
		XIndex:      x,
		YStartIndex: req.YStartIndex,
		YEndIndex:   req.YEndIndex,
		DropKey:     dataUrl,
	}
	return nil
}

// VerifyDrop 验证下拉值是否超限(超过250需要建隐藏sheet)
func VerifyDrop(slice []string) int {
	var (
		vd  int
		str string
	)
	for _, s := range slice {
		str += s
	}
	if tool.LenChar(str) >= 250 {
		vd = 1
	} else {
		vd = 2
	}
	return vd
}
