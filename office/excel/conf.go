package excel

//type ExcelTempTag string
//
//const (
//	DemoTag                               ExcelTempTag = "demo"
//	StudentTag                            ExcelTempTag = "student"
//	TicketSocialProvidentFoundIncreaseTag ExcelTempTag = "TicketSocialProvidentFoundIncreaseTag"
//	TicketSocialIncreaseTag               ExcelTempTag = "TicketSocialIncreaseTag"
//	TicketProvidentFoundIncreaseTag       ExcelTempTag = "TicketProvidentFoundIncreaseTag"
//)
//
//// 模板配置
//var tempTagM = map[ExcelTempTag]string{
//	DemoTag:                               "/file/demo.xlsx",
//	StudentTag:                            "/file/student.xlsx",
//	TicketSocialProvidentFoundIncreaseTag: "/file/社保公积金增员工单-全.xlsx",
//	TicketSocialIncreaseTag:               "/file/社保公积金增员工单-社保.xlsx",
//	TicketProvidentFoundIncreaseTag:       "/file/社保公积金增员工单-公积金.xlsx",
//}
//
//// GetTemp 获取模板文件地址
//func GetTemp(tag string) (string, error) {
//	if _, exi := tempTagM[ExcelTempTag(tag)]; !exi {
//		return "", errors.New("未匹配到模板文件")
//	}
//	return tempTagM[ExcelTempTag(tag)], nil
//}
