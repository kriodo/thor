package header

import (
	"encoding/json"
	"sort"
)

// ParseResult 处理结果
type ParseResult struct {
	ID             string                    // ID
	Status         ProgressStatus            // 状态
	Total          int                       // 数据总数量
	Processed      int                       // 已经处理数
	ImportHeader   []string                  // 上传的excel表头
	ErrMap         map[int][]string          // 错误err Map
	WarnMap        map[int][]string          // 警告warn Map
	ParseContent   []interface{}             // 结果
	DynamicContent map[int]map[string]string // 动态结果
}

// ProgressStatus 状态
type ProgressStatus int

const (
	PROCESSING ProgressStatus = iota + 1 // 进行中
	PROCESSED                            // 已完成
)

// ErrMessage 数据验证错误格式
type ErrMessage struct {
	Line   int      `json:"line"`    // 行号
	ErrMsg []string `json:"err_msg"` // 错误集合
	ExtMsg string   `json:"ext_msg"` // 扩展字段
}

// HasError 是否有错误
func (r *ParseResult) HasError() (map[int][]string, bool) {
	return r.ErrMap, len(r.ErrMap) != 0
}

// Format 格式化数据
func (r *ParseResult) Format(arr interface{}) error {
	marshal, err := json.Marshal(r.ParseContent)
	if err != nil {
		return err
	}
	return json.Unmarshal(marshal, &arr)
}

// 增加错误
func (r *ParseResult) AddError(line int, es ...string) {
	if _, ok := r.ErrMap[line]; !ok {
		r.ErrMap[line] = make([]string, 0)
	}
	r.ErrMap[line] = append(r.ErrMap[line], es...)
}

func (r *ParseResult) AddWarn(line int, es ...string) {
	if _, ok := r.WarnMap[line]; !ok {
		r.WarnMap[line] = make([]string, 0)
	}
	r.WarnMap[line] = append(r.WarnMap[line], es...)
}

// Format 格式化数据
func (r *ParseResult) ErrResult() []*ErrMessage {
	var (
		l    = len(r.ErrMap)
		keys = make([]int, 0, l)
		ret  = make([]*ErrMessage, 0, l)
	)

	for key := range r.ErrMap {
		keys = append(keys, key)
	}
	sort.Ints(keys)
	for _, key := range keys {
		if _, exi := r.ErrMap[key]; exi {
			ret = append(ret, &ErrMessage{
				Line:   key,
				ErrMsg: r.ErrMap[key],
			})
		}
	}
	return ret
}

func (r *ParseResult) WarnResult() []*ErrMessage {
	var (
		l    = len(r.WarnMap)
		keys = make([]int, 0, l)
		ret  = make([]*ErrMessage, 0, l)
	)

	for key := range r.WarnMap {
		keys = append(keys, key)
	}
	sort.Ints(keys)
	for _, key := range keys {
		if _, exi := r.WarnMap[key]; exi {
			ret = append(ret, &ErrMessage{
				Line:   key,
				ErrMsg: r.WarnMap[key],
			})
		}
	}
	return ret
}

// 缓存错误
func (pr *ParseResult) Cache() error {
	r := &Result{
		ID:        pr.ID,
		Status:    pr.Status,
		Total:     pr.Total,
		Processed: pr.Processed,
		ErrList:   pr.ErrResult(),
		WarnList:  pr.WarnResult(),
	}
	_, err := json.Marshal(r)
	if err != nil {
		return err
	}

	return nil
}

// Add 缓存结果增加内容
func (pr *ParseResult) Add() error {
	pr.Processed += 1
	cr := &Result{
		ID:        pr.ID,
		Status:    pr.Status,
		Total:     pr.Total,
		Processed: pr.Processed,
		ErrList:   pr.ErrResult(),
	}
	_, err := json.Marshal(cr)
	if err != nil {
		return err
	}

	return nil
}

// Result 结果
type Result struct {
	ID        string         `json:"id"`        // ID
	Status    ProgressStatus `json:"status"`    // 状态
	Total     int            `json:"total"`     // 总数据量
	Processed int            `json:"processed"` // 已经处理数量
	Percent   uint32         `json:"percent"`   // 进度：0-100
	ErrList   []*ErrMessage  `json:"err_list"`  // 错误err list
	WarnList  []*ErrMessage  `json:"warn_list"` // 警告warn list
}

// 获取结果
func GetResult(id string) (*Result, error) {
	ret := &Result{}

	// 获取百分比
	ret.Percent = ret.getPercent()

	return ret, nil
}

// 获取进度百分比
func (r *Result) getPercent() uint32 {
	if r.Processed >= r.Total {
		return 100
	}
	return uint32(float32(r.Processed) / float32(r.Total) * 100)
}
