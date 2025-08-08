package importer

//type Student struct {
//	ID            int64  `json:"id" excel:"name(序号)"`
//	StudentNumber int64  `json:"student_number" excel:"name(学号);unique(true)"`
//	Name          string `json:"name" excel:"name(姓名);unique(true)"`
//	Sex           int32  `json:"sex" excel:"name(性别);enum(男:1,女:2)"`
//	Age           int32  `json:"age" excel:"name(年龄)"`
//	Birthday      string `json:"birthday" excel:"name(出生日期);date(2006-01-02)"`
//}
//
//func Test_Demo(t *testing.T) {
//	c, err := redis.Dial("tcp", "localhost:6379", redis.DialPassword(""))
//	if err != nil {
//		t.Log(err)
//		return
//	}
//	defer c.Close()
//
//	// new一个excel导入处理器
//	p, err := NewProcessor(&Settings{
//		UniqueID:  "9527", // id
//		RedisConn: c,      // Redis
//		// FileDir:      "./file/upload_student.xlsx", // 上传文件地址
//		OpenValidRow: true,  // 是否开启行验证
//		MaxErrNum:    100,   // 最大错误
//		MaxRowNum:    10000, // 最大数据量
//		// TempHeaderTag: "",                           // 模板表头标识 (默认值=提取结构体name标签的值)
//	}, &Student{})
//	if err != nil {
//		t.Log(err)
//		return
//	}
//
//	// 获取到excel解析内容
//	content, err := p.ParseContent() // 解析的内容
//	if err != nil {
//		t.Log(err)
//		return
//	}
//
//	// 获取错误内容（数据验证错误）
//	errList := content.ErrResult()
//	for _, message := range errList {
//		t.Logf("行号：%d | 错误内容：%+v", message.Line, message.ErrMsg)
//	}
//
//	// 将解析到的内容转换为我们定义的结构体切片
//	ret := make([]*Student, 0, len(content.ParseContent))
//	err = content.Format(&ret) // 数据转换
//	if err != nil {
//		t.Log(err)
//		return
//	}
//	for _, v := range ret {
//		t.Logf("%+v", v)
//	}
//}
//
//func Test_GetResult(t *testing.T) {
//	c, err := redis.Dial("tcp", "localhost:6379", redis.DialPassword(""))
//	if err != nil {
//		t.Log(err)
//		return
//	}
//	defer c.Close()
//
//	result, err := GetResult(c, "9527")
//	if err != nil {
//		t.Log(err)
//		return
//	}
//	j, _ := json.Marshal(result) // nolint
//	t.Logf("处理结果：%s", string(j))
//}
//
//// ValidationRow 自定义的验证数据
//func (p *Student) ValidationRow(r *ParseResult) error {
//	ret := make([]*Student, 0, len(r.ParseContent))
//	err := r.Format(&ret) // 数据转换
//	if err != nil {
//		return err
//	}
//	for k, v := range ret {
//		if v.Name == "张三" {
//			r.AddError(k+1, "他是法外狂徒张三，快报警！！！")
//		}
//	}
//	return nil
//}
