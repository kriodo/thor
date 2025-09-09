package importer

import "github.com/kriodo/thor/office/excel/header"

// ImportSheet 导入sheet相关数据
type ImportSheet struct {
	isSet        bool                         // 是否设置过表头信息
	sheetName    string                       // 表头tree数据
	noHeader     bool                         // 是否有表头
	headerTree   []*header.Header             // 表头tree数据
	fieldInfos   []*header.FieldInfo          // 字段数据 list
	fieldInfoMap map[string]*header.FieldInfo // 字段数据 map （key:info）

	// 表头参数
	headerLength    uint // 表头层数
	headerStartLine int  // 表头起始行号:从0开始
	rowStartLine    int  // 数据起始行号
	maxOrigLen      int  // 数据最大数量

	// 数据
	origRows         [][]string                   // 原始数据
	headerTitleInfos map[string]*ImportHeaderInfo // 表头索引对应表头拼接: 0|"A"|"基本信息.姓名"
	fieldCheckMap    map[string]bool              // 验证表头的字段必存在
	field2dataMap    []map[string]string          // 绑定后数据

}

type ImportHeaderInfo struct {
	Index    int    `json:"index"`     // 下标索引
	Letter   string `json:"letter"`    // 字母索引
	Title    string `json:"title"`     // 标题（层级拼接）
	FieldKey string `json:"field_key"` // 字段key
}
