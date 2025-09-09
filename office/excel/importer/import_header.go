package importer

import (
	"fmt"
	"github.com/kriodo/thor/office/excel/header"
)

// SetHeaderTree 设置list数据 [此模式id、pid必须填写]
func (ir *Importer) SetHeaderTree(tree []*header.Header) *Importer {
	if ir.err != nil {
		return ir
	}
	if ir.cur.isSet {
		ir.err = fmt.Errorf("该sheet重复设置表头信息: %s", ir.cur.sheetName)
		return ir
	}
	if !ir.cur.noHeader { // 设置了表头
		if len(tree) == 0 {
			ir.err = fmt.Errorf("表头不能设置为空: %s", ir.cur.sheetName)
			return ir
		}
		formatData := &header.FormatHeaderData{
			Scene: header.Import,
		}
		ir.err = header.FormatHeaderInfo(formatData, tree, 1, nil)
		if ir.err != nil {
			return ir
		}
		ir.cur.fieldInfos = formatData.FieldInfo
		ir.cur.headerTree = tree
		ir.cur.headerLength = header.MaxLevel(tree, 1)
	}
	ir.cur.isSet = true
	ir.err = ir.handleRowBefore()
	if ir.err != nil {
		return ir
	}
	return ir
}

// SetHeaderListById 设置list数据 [此模式id、pid必须填写]
func (ir *Importer) SetHeaderListById(list []*header.Header) *Importer {
	if ir.err != nil {
		return ir
	}
	err := header.CheckHeaderId(list)
	if err != nil {
		ir.err = err
		return ir
	}
	// 将list数据转为tree数据
	tree := header.ListToTree(list, 0)
	ir.SetHeaderTree(tree)
	return ir
}

// SetHeaderListByPkey 设置list数据 [此模式pkey必须填写]
func (ir *Importer) SetHeaderListByPkey(headers []*header.Header) *Importer {
	var (
		newHeaders  []*header.Header
		headerIdMap = make(map[string]int64)
		id          int64
	)
	for _, head := range headers {
		id++
		head.Id = id
		headerIdMap[head.FieldKey] = id
	}
	for _, head := range headers {
		head.Pid = headerIdMap[head.Pkey]
		newHeaders = append(newHeaders, head)
	}
	return ir.SetHeaderListById(newHeaders)
}

// SetNoHeader 不设置表头
func (ir *Importer) SetNoHeader(startLine, endLine uint) *Importer {
	if startLine > endLine {
		ir.err = fmt.Errorf("开始行号不能大于结束行号")
	}
	ir.cur.noHeader = true
	ir.cur.headerStartLine = int(startLine)
	ir.cur.headerLength = endLine - startLine
	return ir.SetHeaderTree(nil)
}
