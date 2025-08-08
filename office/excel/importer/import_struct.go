package importer

// 导入

//const (
//	ExcelTag = "excel" // excel tag
//
//	ExcelKeySign           = "key"            // 字段key标签
//	ExcelCommentSign       = "comment"        // 字段描述标签
//	ExcelEnumSign          = "enum"           // 枚举标签
//	ExcelUniqueSign        = "unique"         // 唯一标签
//	ExcelDateSign          = "date"           // 日期标签
//	ExcelZhTitleSign       = "zh_title"       // 表头中文描述
//	ExcelOriginalvalueSign = "original_value" // 原始值
//)
//
//var validationSigns = []string{
//	ExcelCommentSign,
//	ExcelEnumSign,
//	ExcelUniqueSign,
//	ExcelDateSign,
//}
//
//// Settings 配置
//type Settings struct {
//	UniqueID                string          // 结果ID
//	OpenValidRow            bool            // 开启自定义行验证
//	OpenValidRowWithContext bool            // 开启自定义行验证(带上下文)
//	Context                 context.Context // 上下文
//	MaxErrNum               int             // 最大错误数量
//	MaxRowNum               int             // 最大数据量
//	NowSheetName               string          // 工作表名称
//	TempHeaderFile          *excelize.File  // 模板file
//	File                    *excelize.File  // 上传文件file [*必填]
//	IsComplexHeader         bool            // 是否复杂表头（复杂表头必填true）
//	NotVerifyHeader         bool            // 不验证表头
//	IsFirstSheetName        bool            // 是否默认直接使用第一个sheet
//	// FileDir        string         // 上传文件地址 [*必填]
//	// TempHeaderTag  string         // 模板表头标签
//}
//
//// Processor 表格上传对象
//type Processor struct {
//	// params
//	uniqueID                string          // 结果ID
//	notVerifyHeader         bool            // 不验证表头
//	tempHeader              [][]string      // 模板表头
//	maxErrNum               int             // 最大错误数量（大致数值，未做严格要求）
//	maxRowNum               int             // 最大数据量
//	sheetName               string          // 表格sheet名称
//	openValidRow            bool            // 自定义行验证
//	openValidRowWithContext bool            // 自定义行验证(带上下文)
//	context                 context.Context // 上下文
//	headerLength            int             // 表头层数
//	rowStartLine            int             // 数据起始行号
//	tempHeaderFile          *excelize.File  // 模板文件file
//	file                    *excelize.File  // 上传文件file
//	isFirstSheetName        bool            // 是否默认直接使用第一个sheet
//
//	// parse content
//	fieldMapping map[string]map[string]string
//	noMapping    bool
//	fieldSlice   []map[string]string
//	appendFiled  []map[string]string
//	body         interface{}
//	val          reflect.Value
//	uniqueMap    map[int][]string
//}
//
//// NewProcessor Excel导入注册
//func NewProcessor(st *Settings, body interface{}) (*Processor, error) {
//	// 设置默认值
//	if st.MaxRowNum <= 0 {
//		st.MaxRowNum = constant.MaxRowNum
//	}
//	if st.MaxErrNum <= 0 {
//		st.MaxErrNum = constant.MaxErrNum
//	}
//	// 如果读取模板文件，并且没有设置sheet名称，并且指定第一个sheet，兼容处理【取第一个名称'因为第一个名可能不是Sheet1'】
//	if st.IsFirstSheetName && st.TempHeaderFile != nil && st.NowSheetName == "" {
//		sheetNameList := st.TempHeaderFile.GetSheetList()
//		if len(sheetNameList) > 0 {
//			st.NowSheetName = sheetNameList[0]
//		}
//	}
//	if st.NowSheetName == "" {
//		st.NowSheetName = constant.DefaultSheetName
//	}
//
//	var tempHeader [][]string
//	// 获取模板表头
//	if st.TempHeaderFile != nil {
//		var err error
//		tempHeader, err = tool.GetTempHeader(st.TempHeaderFile, st.NowSheetName)
//		if err != nil {
//			return nil, err
//		}
//	}
//	if st.UniqueID == "" {
//		st.UniqueID = tool.GetUUID()
//	}
//
//	p := &Processor{
//		uniqueID:                st.UniqueID,
//		notVerifyHeader:         st.NotVerifyHeader,
//		tempHeader:              tempHeader,
//		maxErrNum:               st.MaxErrNum,
//		maxRowNum:               st.MaxRowNum,
//		sheetName:               st.NowSheetName,
//		openValidRow:            st.OpenValidRow,
//		openValidRowWithContext: st.OpenValidRowWithContext,
//		context:                 st.Context,
//		headerLength:            len(tempHeader),
//		rowStartLine:            len(tempHeader),
//		tempHeaderFile:          st.TempHeaderFile,
//		file:                    st.File,
//		isFirstSheetName:        st.IsFirstSheetName,
//		fieldMapping:            make(map[string]map[string]string),
//		noMapping:               st.IsComplexHeader,
//		fieldSlice:              make([]map[string]string, 0),
//		appendFiled:             nil,
//		body:                    body,
//		val:                     reflect.Value{},
//		uniqueMap:               nil,
//	}
//	p.val = reflect.ValueOf(body)
//	if p.val.Kind() != reflect.Ptr {
//		return nil, errors.New("body must be pointer struct")
//	}
//	// 生成结构体与Excel头映射关系
//	p.generateMapping(p.val, "")
//
//	return p, nil
//}
//
//// WithTempHeader 添加自定义表头
//func (p *Processor) WithTempHeader(header []map[string]string) {
//	// p.fieldSlice = append(p.fieldSlice, header...)
//	p.appendFiled = header
//}
//func (p *Processor) SetRowStartLine(line int) {
//	p.rowStartLine = line
//}
//
//// Mapping 接口
//type Mapping interface {
//	ValidationRow(r *ParseResult) error // 验证行数据 [根据需求自己去实现它]
//}
//
//// Mapping2 接口
//type Mapping2 interface {
//	ValidationRowWithContext(ctx context.Context, r *ParseResult) error // 验证行数据 [根据需求自己去实现它]
//}
//
//// ValidationRow 验证数据
//func (p *Processor) ValidationRow(r *ParseResult) error {
//	return nil
//}
//
//// ValidationRow 验证数据
//func (p *Processor) ValidationRowWithContext(ctx context.Context, r *ParseResult) error {
//	return nil
//}
//
//// ParseContent 解析内容
//func (p *Processor) ParseContent() (*ParseResult, error) {
//	var (
//		err    error
//		header [][]string
//		rows   [][]string
//	)
//
//	defer func() {
//		if err != nil {
//			// 设置缓存结果
//			ret := ParseResult{
//				ID:     p.uniqueID,
//				Status: PROCESSED,
//				ErrMap: make(map[int][]string),
//			}
//			ret.AddError(-1, err.Error())
//			err = ret.Cache()
//			if err != nil {
//				fmt.Printf("解析错误：%s", err.Error())
//			}
//		}
//	}()
//	// 参数验证
//	if p.maxRowNum < 1 && p.headerLength < 1 && p.maxErrNum < 1 {
//		err = errors.New("配置参数错误")
//		return nil, err
//	}
//	if len(p.tempHeader) <= 0 {
//		err = errors.New("未匹配到模板文件")
//		return nil, err
//	}
//	if p.file == nil {
//		err = errors.New("未上传文件")
//		return nil, err
//	}
//
//	rows, header, err = p.getUploadExcel()
//	if err != nil {
//		return nil, err
//	}
//	// 验证表头
//	if !p.notVerifyHeader {
//		err = verifyHeader(p.tempHeader, header)
//		if err != nil {
//			return nil, err
//		}
//	}
//
//	return p.rows(rows, header)
//}
//
//// 生成结构体与Excel头映射关系
//func (p *Processor) generateMapping(val reflect.Value, baseField string) {
//	switch val.Kind() { // nolint
//	case reflect.Struct:
//	case reflect.Ptr:
//		// 当结构体指针或字段指针为空，则创建一个指针指向
//		if val.IsNil() {
//			newValue := reflect.New(val.Type().Elem())
//			val = reflect.NewAt(val.Type().Elem(), unsafe.Pointer(newValue.Pointer()))
//		}
//		val = val.Elem()
//		p.generateMapping(val, baseField)
//		return
//	default:
//		return
//	}
//
//	var (
//		typ    = val.Type()
//		header []string
//	)
//	for i := 0; i < val.NumField(); i++ {
//		fieldName := typ.Field(i).Name
//		if baseField != "" {
//			fieldName = fmt.Sprintf("%s.%s", baseField, fieldName)
//		}
//		excel, ok := typ.Field(i).Tag.Lookup(ExcelTag)
//		if !ok {
//			// 生成嵌套结构体的映射关系
//			fieldVal := val.Field(i)
//			p.generateMapping(fieldVal, fieldName)
//			continue
//		}
//		var (
//			mappingName string
//			m           = make(map[string]string, len(validationSigns))
//		)
//		for _, tag := range validationSigns {
//			if tag == ExcelCommentSign {
//				m[tag] = fieldName
//				mappingName, _ = stringMatchExport(excel, regexp.MustCompile(fmt.Sprintf(`%s\((.*?)\)`, ExcelCommentSign))) // nolint
//				header = append(header, mappingName)
//				continue
//			}
//			m[tag], _ = stringMatchExport(excel, regexp.MustCompile(fmt.Sprintf(`%s\((.*?)\)`, tag))) // nolint
//		}
//		key := fmt.Sprintf("%d_%s", i, strings.TrimSpace(mappingName))
//		m[ExcelZhTitleSign] = strings.TrimSpace(mappingName)
//		p.fieldMapping[key] = m
//		p.fieldSlice = append(p.fieldSlice, m)
//	}
//	if p.tempHeaderFile == nil {
//		p.tempHeader = append(p.tempHeader, header)
//		p.headerLength = 1
//	}
//}
//
//// 数据绑定
//func (p *Processor) rows(rows, header [][]string) (*ParseResult, error) {
//	var (
//		ret = &ParseResult{
//			ID:     p.uniqueID,
//			Total:  len(rows),
//			Status: PROCESSING,
//			ErrMap: make(map[int][]string),
//		}
//		count   int
//		uniqueM = make(map[string][]string)
//		err     error
//	)
//	// 设置缓存结果
//	err = ret.Cache()
//	if err != nil {
//		return nil, err
//	}
//	if len(p.appendFiled) > 0 {
//		startIndex := len(p.fieldSlice)
//		ret.DynamicContent = make(map[int]map[string]string)
//		for k1 := range rows {
//			m := make(map[string]string)
//			for k2, filed := range p.appendFiled {
//				var colVal string
//				if len(rows[k1]) > startIndex+k2 {
//					colVal = strings.TrimSpace(rows[k1][startIndex+k2])
//				}
//				m[filed[ExcelKeySign]] = colVal // 原始数据
//			}
//			ret.DynamicContent[k1] = m
//		}
//	}
//
//	if p.noMapping { // 顺序赋值（复杂表头）
//		for i := 0; i < len(rows); i++ {
//			var (
//				dateErrList       []string
//				mappingErrList    []string
//				parseValueErrList []string
//				line              = p.headerLength + i + p.headerLength // 行号
//				newBodyVal        = reflect.New(p.val.Type().Elem())
//				uniqueKey         string
//			)
//			newBodyVal.Elem().Set(p.val.Elem())
//			// 循环行数据
//			for index, m := range p.fieldSlice {
//				if len(rows[i]) < index+1 {
//					continue
//				}
//				var (
//					colVal        = strings.TrimSpace(rows[i][index])
//					key           string
//					mappingHeader string
//				)
//				mappingHeader = m[ExcelZhTitleSign]
//				// 去除列的前后空格
//				key = fmt.Sprintf("%d_%s", index, strings.TrimSpace(mappingHeader))
//				tagMap, ok := p.fieldMapping[key]
//				if !ok {
//					continue
//				}
//				// 判断是否有列唯一标签
//				uniqueKey = p.uniqueFormat(uniqueKey, colVal, tagMap)
//				// 格式化时间
//				dateErrList = p.dateFormat(&colVal, tagMap)
//				if len(dateErrList) > 0 {
//					ret.AddError(line, dateErrList...)
//				}
//				// 值映射转换
//				mappingErrList = p.mappingFormat(mappingHeader, &colVal, tagMap)
//				if len(mappingErrList) > 0 {
//					ret.AddError(line, mappingErrList...)
//					continue
//				}
//				// 参数赋值
//				parseValueErrList, err = p.parseValue(newBodyVal, tagMap[ExcelCommentSign], mappingHeader, colVal)
//				if err != nil {
//					ret.AddError(line, "参数赋值错误")
//					continue
//				}
//				if len(parseValueErrList) > 0 {
//					ret.ErrMap[line] = append(ret.ErrMap[line], parseValueErrList...)
//				}
//				count += len(dateErrList) + len(mappingErrList) + len(parseValueErrList)
//				// 达到最大错误，无需再验证下去了
//				if count >= p.maxErrNum {
//					return ret, nil
//				}
//			}
//			p.body = newBodyVal.Interface()
//			ret.ParseContent = append(ret.ParseContent, p.body)
//			// 列唯一性校验
//			if uniqueKey != "" {
//				uniqueM[uniqueKey] = append(uniqueM[uniqueKey], strconv.Itoa(line))
//			}
//
//			// 缓存结果计数+1
//			err = ret.Add()
//			if err != nil {
//				return nil, err
//			}
//		}
//	} else { // 表头赋值
//		for i := 0; i < len(rows); i++ {
//			var (
//				dateErrList       []string
//				mappingErrList    []string
//				parseValueErrList []string
//				line              = p.headerLength + i + p.headerLength // 行号
//				newBodyVal        = reflect.New(p.val.Type().Elem())
//				uniqueKey         string
//			)
//			newBodyVal.Elem().Set(p.val.Elem())
//			// 循环行数据
//			for index, col := range rows[i] {
//				var (
//					colVal        = strings.TrimSpace(col)
//					key           string
//					mappingHeader string
//				)
//				mappingHeader = header[p.headerLength-1][index]
//				// 去除列的前后空格
//				key = fmt.Sprintf("%d_%s", index, strings.TrimSpace(mappingHeader))
//				tagMap, ok := p.fieldMapping[key]
//				if !ok {
//					continue
//				}
//				// 判断是否有列唯一标签
//				uniqueKey = p.uniqueFormat(uniqueKey, colVal, tagMap)
//				// 格式化时间
//				dateErrList = p.dateFormat(&colVal, tagMap)
//				if len(dateErrList) > 0 {
//					ret.AddError(line, dateErrList...)
//				}
//				// 值映射转换
//				mappingErrList = p.mappingFormat(mappingHeader, &colVal, tagMap)
//				if len(mappingErrList) > 0 {
//					ret.AddError(line, mappingErrList...)
//					continue
//				}
//				// 参数赋值
//				parseValueErrList, err = p.parseValue(newBodyVal, tagMap[ExcelCommentSign], mappingHeader, colVal)
//				if err != nil {
//					ret.AddError(line, "参数赋值错误")
//					continue
//				}
//				if len(parseValueErrList) > 0 {
//					ret.ErrMap[line] = append(ret.ErrMap[line], parseValueErrList...)
//				}
//				count += len(dateErrList) + len(mappingErrList) + len(parseValueErrList)
//				// 达到最大错误，无需再验证下去了
//				if count >= p.maxErrNum {
//					return ret, nil
//				}
//			}
//			p.body = newBodyVal.Interface()
//			ret.ParseContent = append(ret.ParseContent, p.body)
//			// 列唯一性校验
//			if uniqueKey != "" {
//				uniqueM[uniqueKey] = append(uniqueM[uniqueKey], strconv.Itoa(line))
//			}
//
//			// 缓存结果计数+1
//			err = ret.Add()
//			if err != nil {
//				return nil, err
//			}
//		}
//	}
//
//	// 列唯一性校验
//	if len(uniqueM) > 0 {
//		for _, repeat := range uniqueM {
//			if len(repeat) <= 1 {
//				continue
//			}
//			k, _ := strconv.Atoi(repeat[0]) // nolint
//			ret.AddError(k, fmt.Sprintf("第%s行数据重复", strings.Join(repeat, "、")))
//		}
//	}
//	// 自定义参数验证
//	err = p.definedValidRow(ret)
//	if err != nil {
//		return nil, err
//	}
//
//	defer func() {
//		if err != nil {
//			ret.AddError(-1, err.Error())
//		}
//		ret.Status = PROCESSED
//		err = ret.Cache()
//		if err != nil {
//			return
//		}
//	}()
//
//	return ret, nil
//}
//
//// 自定义行验证
//func (p *Processor) definedValidRow(ret *ParseResult) error {
//	if p.openValidRow {
//		inf, ok := p.body.(Mapping)
//		if ok {
//			// 如果未实现自定义验证接口，调用默认验证
//			return inf.ValidationRow(ret)
//		}
//	}
//	if p.openValidRowWithContext && p.context != nil {
//		inf, ok := p.body.(Mapping2)
//		if ok {
//			// 如果未实现自定义验证接口，调用默认验证
//			return inf.ValidationRowWithContext(p.context, ret)
//		}
//	}
//	return nil
//}
//
//// 获取处理数据的excel数据
//func (p *Processor) getUploadExcel() ([][]string, [][]string, error) {
//	if p.isFirstSheetName {
//		sheetNameList := p.file.GetSheetList()
//		if len(sheetNameList) > 0 {
//			p.sheetName = sheetNameList[0]
//		}
//	}
//	// 获取处理excel数据
//	rows, err := readExcel2(p.file, p.sheetName)
//	if err != nil {
//		return nil, nil, err
//	}
//	l := len(rows)
//	if l <= 0 {
//		return nil, nil, errors.New("文件有效数据为空")
//	}
//	if p.headerLength > l {
//		return nil, nil, errors.New("表格表头错误，请使用正确的表格")
//	}
//	// excel数据行数限制
//	if l > p.maxRowNum {
//		return nil, nil, fmt.Errorf("文件数据量超限：%d", p.maxErrNum)
//	}
//	if p.rowStartLine > len(rows) {
//		return nil, nil, fmt.Errorf("数据起始行号错误：%d", p.rowStartLine)
//	}
//	header := rows[:p.headerLength]
//	data := rows[p.rowStartLine:]
//
//	return data, header, nil
//}
//
//// 值映射转换
//func (p *Processor) mappingFormat(mappingHeader string, col *string, mappingField map[string]string) []string {
//	errList := make([]string, 0)
//	format, ok := mappingField[ExcelEnumSign]
//	if !ok || format == "" {
//		return errList
//	}
//	mappingValues := make(map[string]string)
//	formatStr := strings.Split(format, ",")
//	for _, format := range formatStr {
//		n := strings.SplitN(format, ":", 2)
//		if len(n) != 2 {
//			continue
//		}
//		mappingValues[n[0]] = n[1]
//	}
//	val, ok := mappingValues[*col]
//	if ok {
//		*col = val
//		return errList
//	}
//	errList = append(errList, fmt.Sprintf("%s单元格存在非法输入", mappingHeader))
//	return errList
//}
//
//// 参数赋值
//func (p *Processor) parseValue(val reflect.Value, fieldAddr, mappingHeader, col string) ([]string, error) {
//	errList := make([]string, 0)
//	fields := strings.Split(fieldAddr, ".")
//	if len(fields) == 0 {
//		return errList, nil
//	}
//	for _, field := range fields {
//		if val.Kind() == reflect.Ptr {
//			val = val.Elem()
//		}
//		val = val.FieldByName(field)
//		errs, err := p.parse(val, col, mappingHeader)
//		if err != nil {
//			return errList, err
//		}
//		errList = append(errList, errs...)
//	}
//	return errList, nil
//}
//
//// 解析
//func (p *Processor) parse(val reflect.Value, col, mappingHeader string) ([]string, error) {
//	errList := make([]string, 0)
//	var err error
//	switch val.Kind() { // nolint
//	case reflect.String:
//		val.SetString(col)
//	case reflect.Bool:
//		parseBool, e := strconv.ParseBool(col)
//		if e != nil {
//			errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非bool类型值", mappingHeader))
//		}
//		val.SetBool(parseBool)
//	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
//		var value int64
//		if col != "" {
//			value, err = strconv.ParseInt(col, 10, 64)
//			if err != nil {
//				errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非整形数值", mappingHeader))
//			}
//		}
//		val.SetInt(value)
//	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
//		var value uint64
//		if col != "" {
//			value, err = strconv.ParseUint(col, 10, 64)
//			if err != nil {
//				errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非整形数值", mappingHeader))
//			}
//		}
//		val.SetUint(value)
//	case reflect.Float32, reflect.Float64:
//		var value float64
//		if col != "" {
//			value, err = strconv.ParseFloat(col, 64)
//			if err != nil {
//				errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非浮点型数值", mappingHeader))
//			}
//		}
//		val.SetFloat(value)
//	case reflect.Struct:
//		return errList, nil
//	case reflect.Ptr:
//		// 初始化指针
//		value := reflect.New(val.Type().Elem())
//		val.Set(value)
//		var errs []string
//		errs, err = p.parse(val.Elem(), col, mappingHeader)
//		if err != nil {
//			break
//		}
//		errList = append(errList, errs...)
//	default:
//		return errList, fmt.Errorf("excel column[%s] parseValue unsupported type[%v] mappings", mappingHeader, val.Kind().String())
//	}
//	return errList, nil
//}
