package exporter

import (
	"github.com/kriodo/thor/office/excel/header"
	"os"
	"testing"
)

func TestExporter(t *testing.T) {
	var (
		fileName = "../test/导出测试.xlsx"
	)
	os.Remove(fileName)
	export, err := NewExporter("测试-1")
	if err != nil {
		t.Log(err)
		return
	}
	// 设置表头从第2列开始
	export.SetHeaderStartX(2)
	// 设置表头从第3行开始
	export.SetHeaderStartY(3)
	// tree结构的表头
	treeHeader := []*header.Header{
		{Title: "姓名", FieldKey: "user_name"},
		{Title: "户籍", FieldKey: "household_registration"},
		{Title: "户口所在城市", FieldKey: "household_registration_city"},
		{Title: "工作城市", FieldKey: "work_city"},
		{Title: "子项目名称", FieldKey: "subitem_name"},
		{Title: "调整前", FieldKey: "before", Children: []*header.Header{
			{Title: "养老", FieldKey: "before_old_age", Children: []*header.Header{
				{Title: "企业基数", FieldKey: "before_old_age_company_base"},
				{Title: "企业比例", FieldKey: "before_old_age_company_rate"},
				{Title: "个人基数", FieldKey: "before_old_age_personnel_base"},
				{Title: "个人比例", FieldKey: "before_old_age_personnel_rate"},
			}}}},
		{Title: "调整后", FieldKey: "after", Children: []*header.Header{
			{Title: "养老", FieldKey: "after_old_age", Children: []*header.Header{
				{Title: "企业基数", FieldKey: "after_old_age_company_base"},
				{Title: "企业比例", FieldKey: "after_old_age_company_rate"},
				{Title: "个人基数", FieldKey: "after_old_age_personnel_base"},
				{Title: "个人比例", FieldKey: "after_old_age_personnel_rate"},
			}}}},
		{Title: "备注", FieldKey: "remark"},
	}
	err = export.SetTree(treeHeader).Error()
	if err != nil {
		t.Log(err)
		return
	}
	// 设置下拉框
	err = export.SetDropByFieldKey([]*FieldDropInfo{
		{
			UniqueKey: "户籍",
			FieldKeys: []string{"household_registration"},
			YEndIndex: 100,
			ValueList: []string{"安徽省", "北京市", "福建省", "甘肃省", "广东省", "广西壮族自治区", "贵州省", "海南省", "河北省", "黑龙江省", "河南省", "湖北省", "湖南省", "吉林省", "江苏省", "江西省", "辽宁省", "内蒙古自治区", "宁夏回族自治区", "青海省", "山东省", "山西省", "陕西省", "上海市", "四川省", "天津市", "西藏自治区", "新疆维吾尔自治区", "云南省", "浙江省", "重庆市", "香港特别行政区", "澳门特别行政区"},
		},
		{
			UniqueKey: "城市",
			FieldKeys: []string{"household_registration_city", "work_city"},
			YEndIndex: 100,
			ValueList: []string{"石家庄市", "唐山市", "秦皇岛市", "邯郸市", "邢台市", "保定市", "张家口市", "承德市", "沧州市", "廊坊市", "衡水市", "太原市", "大同市", "阳泉市", "长治市", "晋城市", "朔州市", "晋中市", "运城市", "忻州市", "临汾市", "吕梁市", "呼和浩特市", "包头市", "乌海市", "赤峰市", "通辽市", "鄂尔多斯市", "呼伦贝尔市", "巴彦淖尔市", "乌兰察布市", "沈阳市", "大连市", "鞍山市", "抚顺市", "本溪市", "丹东市", "锦州市", "营口市", "阜新市", "辽阳市", "盘锦市", "铁岭市", "朝阳市", "葫芦岛市", "长春市", "吉林市", "四平市", "辽源市", "通化市", "白山市", "松原市", "白城市", "哈尔滨市", "齐齐哈尔市", "鸡西市", "鹤岗市", "双鸭山市", "大庆市", "伊春市", "佳木斯市", "七台河市", "牡丹江市", "黑河市", "绥化市", "南京市", "无锡市", "徐州市", "常州市", "苏州市", "南通市", "连云港市", "淮安市", "盐城市", "扬州市", "镇江市", "泰州市", "宿迁市", "杭州市", "宁波市", "温州市", "嘉兴市", "湖州市", "绍兴市", "金华市", "衢州市", "舟山市", "台州市", "丽水市", "合肥市", "芜湖市", "蚌埠市", "淮南市", "马鞍山市", "淮北市", "铜陵市", "安庆市", "黄山市", "滁州市", "阜阳市", "宿州市", "六安市", "亳州市", "池州市", "宣城市", "福州市", "厦门市", "三明市", "莆田市", "泉州市", "漳州市", "南平市", "龙岩市", "宁德市", "南昌市", "景德镇市", "萍乡市", "九江市", "新余市", "鹰潭市", "赣州市", "吉安市", "宜春市", "抚州市", "上饶市", "济南市", "青岛市", "淄博市", "枣庄市", "东营市", "烟台市", "潍坊市", "济宁市", "泰安市", "威海市", "日照市", "临沂市", "德州市", "聊城市", "滨州市", "菏泽市", "郑州市", "开封市", "洛阳市", "平顶山市", "安阳市", "鹤壁市", "新乡市", "焦作市", "濮阳市", "许昌市", "漯河市", "三门峡市", "南阳市", "商丘市", "信阳市", "周口市", "驻马店市", "武汉市", "黄石市", "十堰市", "宜昌市", "襄阳市", "鄂州市", "荆门市", "孝感市", "荆州市", "黄冈市", "咸宁市", "随州市", "长沙市", "株洲市", "湘潭市", "衡阳市", "邵阳市", "岳阳市", "常德市", "张家界市", "益阳市", "郴州市", "永州市", "怀化市", "娄底市", "广州市", "韶关市", "深圳市", "珠海市", "汕头市", "佛山市", "江门市", "湛江市", "茂名市", "肇庆市", "惠州市", "梅州市", "汕尾市", "河源市", "阳江市", "清远市", "东莞市", "中山市", "潮州市", "揭阳市", "云浮市", "南宁市", "柳州市", "桂林市", "梧州市", "北海市", "防城港市", "钦州市", "贵港市", "玉林市", "百色市", "贺州市", "河池市", "来宾市", "崇左市", "海口市", "三亚市", "三沙市", "儋州市", "成都市", "自贡市", "攀枝花市", "泸州市", "德阳市", "绵阳市", "广元市", "遂宁市", "内江市", "乐山市", "南充市", "眉山市", "宜宾市", "广安市", "达州市", "雅安市", "巴中市", "资阳市", "贵阳市", "六盘水市", "遵义市", "安顺市", "毕节市", "铜仁市", "昆明市", "曲靖市", "玉溪市", "保山市", "昭通市", "丽江市", "普洱市", "临沧市", "拉萨市", "日喀则市", "昌都市", "林芝市", "山南市", "那曲市", "西安市", "铜川市", "宝鸡市", "咸阳市", "渭南市", "汉中市", "延安市", "榆林市", "安康市", "商洛市", "兰州市", "嘉峪关市", "金昌市", "白银市", "天水市", "武威市", "张掖市", "平凉市", "酒泉市", "庆阳市", "定西市", "陇南市", "西宁市", "海东市", "银川市", "石嘴山市", "吴忠市", "固原市", "中卫市", "乌鲁木齐市", "克拉玛依市", "吐鲁番市", "哈密市"},
		},
		{
			UniqueKey: "比例下拉",
			FieldKeys: []string{"before_old_age_company_rate", "before_old_age_personnel_rate", "after_old_age_company_rate", "after_old_age_personnel_rate"},
			YEndIndex: 100,
			ValueList: []string{"1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%", "9%", "10%", "11%", "12%"},
		},
	})
	if err != nil {
		t.Log(err)
		return
	}
	export, err = export.AddSheet("测试-2")
	if err != nil {
		t.Log(err)
		return
	}
	// list表头
	var listHeaders []*header.Header
	listHeaders = append(listHeaders, &header.Header{Id: 101, Pid: 0, Title: "子项目名称", FieldKey: "subitem_name"})
	listHeaders = append(listHeaders, &header.Header{Id: 102, Pid: 0, Title: "姓名", Weight: 1, FieldKey: "name"})
	listHeaders = append(listHeaders, &header.Header{Id: 103, Pid: 0, Title: "证件号", FieldKey: "id_card"})
	listHeaders = append(listHeaders, &header.Header{Id: 104, Pid: 0, Title: "工作城市", FieldKey: "work_city"})
	listHeaders = append(listHeaders, &header.Header{Id: 105, Pid: 0, Title: "调整前", FieldKey: "before"})
	listHeaders = append(listHeaders, &header.Header{Id: 1001, Pid: 105, Title: "养老", FieldKey: "before_old_age"})
	listHeaders = append(listHeaders, &header.Header{Id: 10001, Pid: 1001, Title: "企业", FieldKey: "before_old_age_company"})
	listHeaders = append(listHeaders, &header.Header{Id: 100001, Pid: 10001, Title: "企业基数", FieldKey: "before_old_age_company_base"})
	listHeaders = append(listHeaders, &header.Header{Id: 100002, Pid: 10001, Title: "企业比例", Weight: 1, FieldKey: "before_old_age_company_rate"})
	listHeaders = append(listHeaders, &header.Header{Id: 10002, Pid: 1001, Title: "个人", FieldKey: "before_old_age_personnel"})
	listHeaders = append(listHeaders, &header.Header{Id: 100003, Pid: 10002, Title: "个人基数", FieldKey: "before_old_age_personnel_base"})
	listHeaders = append(listHeaders, &header.Header{Id: 100004, Pid: 10002, Title: "个人比例", FieldKey: "before_old_age_personnel_rate"})
	listHeaders = append(listHeaders, &header.Header{Id: 106, Pid: 0, Title: "调整后", FieldKey: "after"})
	listHeaders = append(listHeaders, &header.Header{Id: 1002, Pid: 106, Title: "养老", FieldKey: "after_old_age"})
	listHeaders = append(listHeaders, &header.Header{Id: 10003, Pid: 1002, Title: "企业", FieldKey: "after_old_age_company"})
	listHeaders = append(listHeaders, &header.Header{Id: 100005, Pid: 10003, Title: "企业基数", FieldKey: "after_old_age_company_base"})
	listHeaders = append(listHeaders, &header.Header{Id: 100006, Pid: 10003, Title: "企业比例", FieldKey: "after_old_age_company_rate"})
	listHeaders = append(listHeaders, &header.Header{Id: 10004, Pid: 1002, Title: "个人", FieldKey: "after_old_age_personnel"})
	listHeaders = append(listHeaders, &header.Header{Id: 100007, Pid: 10004, Title: "个人基数", FieldKey: "after_old_age_personnel_base"})
	listHeaders = append(listHeaders, &header.Header{Id: 100008, Pid: 10004, Title: "个人比例", FieldKey: "after_old_age_personnel_rate"})
	listHeaders = append(listHeaders, &header.Header{Id: 107, Pid: 0, Title: "备注", FieldKey: "remark"})
	export.SetListById(listHeaders)
	// 设置下拉框
	err = export.SetDrop([]*DropInfo{
		{
			UniqueKey:   "城市",
			XIndex:      export.GetFieldXIndex("work_city"),
			YStartIndex: export.GetDataStartY(),
			YEndIndex:   100,
		},
		{
			UniqueKey:   "比例下拉",
			XIndex:      export.GetFieldXIndex("before_old_age_company_rate"),
			YStartIndex: export.GetDataStartY(),
			YEndIndex:   100,
			ValueList:   []string{"1%", "2%", "3%", "4%", "5%", "6%", "7%", "8%", "9%", "10%", "11%", "12%"},
		},
	})
	if err != nil {
		t.Log(err)
		return
	}

	export, err = export.AddSheet("测试-3")
	if err != nil {
		t.Log(err)
		return
	}
	// list表头
	var listHeader2s []*header.Header
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "", Title: "子项目名称", FieldKey: "subitem_name"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "", Title: "姓名", FieldKey: "name"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "", Title: "证件号", FieldKey: "id_card"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "", Title: "调整前", FieldKey: "before"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "before", Title: "养老", FieldKey: "before_old_age"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "before_old_age", Title: "企业", FieldKey: "before_old_age_company"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "before_old_age_company", Title: "企业基数", FieldKey: "before_old_age_company_base"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "before_old_age_company", Title: "企业比例", Weight: 1, FieldKey: "before_old_age_company_rate"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "before_old_age", Title: "个人", FieldKey: "before_old_age_personnel"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "before_old_age_personnel", Title: "个人基数", FieldKey: "before_old_age_personnel_base"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "before_old_age_personnel", Title: "个人比例", FieldKey: "before_old_age_personnel_rate"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "", Title: "调整后", FieldKey: "after"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "after", Title: "养老", FieldKey: "after_old_age"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "after_old_age", Title: "企业", FieldKey: "after_old_age_company"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "after_old_age_company", Title: "企业基数", FieldKey: "after_old_age_company_base"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "after_old_age_company", Title: "企业比例", FieldKey: "after_old_age_company_rate"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "after_old_age", Title: "个人", FieldKey: "after_old_age_personnel"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "after_old_age_personnel", Title: "个人基数", FieldKey: "after_old_age_personnel_base"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "after_old_age_personnel", Title: "个人比例", FieldKey: "after_old_age_personnel_rate"})
	listHeader2s = append(listHeader2s, &header.Header{Pkey: "", Title: "备注", FieldKey: "remark"})
	export.SetListById(listHeaders)
	err = export.SaveAs(fileName)
	if err != nil {
		t.Log(err)
		return
	}
	t.Log("------ success")
}
